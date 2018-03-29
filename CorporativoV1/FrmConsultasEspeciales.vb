Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class FrmConsultasEspeciales
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''' ********************************************************************************************************************
    ''' MODIFICACION DE LA CONSULTA DE REPARACIONES - ESTABA MUY LENTA POR UNA FUNCION CONTENIDA EN EL SELECT.  SE MEJORO EL
    ''' FUNCIONAMIENTO OPERATIVO DEL COMBO Y DEL FILTRADO DEL NOMBRE DEL CLIENTE
    ''' AFECTA :
    ''' Administración de Reparaciones - frmCorpoControlReparaciones_Corpo
    ''' Reporte de Reparaciones - frmVtasRptReparaciones_Corpo
    ''' 15SEP2006 - MAVF
    ''' ********************************************************************************************************************

    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents txtNombre As System.Windows.Forms.TextBox
    Public WithEvents chkTodasSucursales As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.Panel
    Public WithEvents _Label1_3 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Public RenAnt As Integer
    Public I As Integer
    Public intCodSucursal As Integer
    Public FueraChange As Boolean
    Public tecla As Integer
    Public strSQL As String
    Public strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
    Public strCaptionForm As String 'Titulo que mostrara el formulario de consultas
    Public strControlActual As String 'Nombre del control actual
    Public strFormaActual As String 'Nombre de la Forma actual
    Public Columna As Integer
    Public cWHERE As String
    Public cQUERYBusqueda As String
    Public WithEvents Flexdet As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public gCodAlmacenInicial As Integer

    Private Sub chkTodasSucursales_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodasSucursales.CheckStateChanged

        If FueraChange Then Exit Sub

        If chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
            FueraChange = True
            dbcSucursales.Text = ""
            dbcSucursales.Refresh()
            txtNombre.Text = ""
            txtNombre.Refresh()
            FueraChange = False
            intCodSucursal = 0
            dbcSucursales.Enabled = False
            Encabezado()
            Buscar()
            txtNombre.Focus()
        Else
            dbcSucursales.Enabled = True
            FueraChange = True
            dbcSucursales.Text = ""
            txtNombre.Text = ""
            FueraChange = False
            Encabezado()
            intCodSucursal = gCodAlmacenInicial
            LlenaDatosSucursal()
            Buscar()
            dbcSucursales.Focus()
        End If

    End Sub

    Private Sub dbcSucursales_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursales" Then
        '    Exit Sub
        'End If
        FueraChange = True
        txtNombre.Text = ""
        FueraChange = False
        gStrSql = "SELECT CodAlmacen,Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE TipoAlmacen ='P' and  DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla, dbcSucursales)
        PonerCodigoSucursal()
        Buscar()
    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen  FROM CatAlmacen where  TipoAlmacen ='P'  ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
    End Sub

    Private Sub dbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkTodasSucursales.Focus()
        End If
    End Sub

    Private Sub dbcSucursales_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyUp

    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE  TipoAlmacen ='P' and  DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
    End Sub

    Private Sub dbcSucursales_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcSucursales.MouseUp
        'PonerCodigoSucursal()
        'Buscar()
    End Sub

    Private Sub FlexDet_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles FlexDet.DblClick
        Aceptar()
    End Sub

    Sub PonerCodigoSucursal()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        If RsGral.RecordCount = 0 Then
            intCodSucursal = 0
        Else
            intCodSucursal = RsGral.Fields("CodAlmacen").Value
        End If

    End Sub

    Private Sub Flexdet_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Flexdet.Enter
        If Flexdet.Rows > 1 Then
            Flexdet.ColSel = 1
            Flexdet.ColSel = 6
            Flexdet.HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightAlways
            Flexdet.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusNone
        End If
    End Sub

    Private Sub FlexDet_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles FlexDet.KeyPressEvent
        If eventArgs.keyAscii = 13 Then
            Aceptar()
        End If
    End Sub

    Private Sub Flexdet_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Flexdet.Leave
        Flexdet.FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
        Flexdet.HighLight = MSHierarchicalFlexGridLib.HighLightSettings.flexHighlightNever
    End Sub

    Private Sub FrmConsultasEspeciales_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Me.Flexdet.Row = 1
        Me.Flexdet.Col = 0
        Me.Flexdet.ColSel = 0
    End Sub

    Public Sub Aceptar()
        On Error GoTo Merr
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        With Flexdet
            If Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)) <> "" Then
                Select Case strTag
                    Case "FRMCORPOCONTROLREPARACIONES_CORPO.TXTFOLIO"
                        With frmCorpoControlReparaciones_Corpo
                            .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                            frmCorpoControlReparaciones_Corpo.bandera = False
                            .LlenaDatos()
                            Me.Close()
                        End With
                    Case Else
                        Me.Close()
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                        Exit Sub
                End Select
                'System.Windows.Forms.SendKeys.Send("{ENTER}")
            End If
        End With
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub
Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MostrarError("Ha ocurrido un error")
    End Sub

    Private Sub FrmConsultasEspeciales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then Me.Close()
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If UCase(Trim(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name)) <> UCase("flexdet") Then ModEstandar.AvanzarTab(Me)
        End If
    End Sub

    Private Sub FrmConsultasEspeciales_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub FrmConsultasEspeciales_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'InitializeComponent()
        KeyPreview = True
        ModEstandar.CentrarForma(Me)
        'System.Windows.Forms.SendKeys.Send("{RIGHT}")
        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        'strFormaActual = UCase(System.Windows.Forms.Form.ActiveForm.Name)
        strTag = UCase(strFormaActual & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control
        gCodAlmacenInicial = intCodSucursal
        FueraChange = True
        chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
        FueraChange = False
        dbcSucursales.Enabled = False
        Buscar()
    End Sub

    Private Sub FrmConsultasEspeciales_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Sub Encabezado()

        With Me.Flexdet
            '''.Clear
            .Rows = 2
            .Rows = 10
            '''.COLS = 7

            '''      .ColWidth(0) = 0
            '''      .ColWidth(1) = 2000  'fOLIO
            '''      .ColWidth(2) = 1200  'fECHA
            '''      .ColWidth(3) = 3800  'CLIENTE
            '''      .ColWidth(4) = 1000  'TIPO VENTA
            '''      .ColWidth(5) = 1000  'TOTAL
            '''      .ColWidth(6) = 2000  'ESTATUS
            '''      .ColAlignment(1) = flexAlignRightCenter
            '''      .ColAlignment(2) = flexAlignCenterCenter
            '''      .ColAlignment(3) = flexAlignLeftCenter
            '''      .ColAlignment(4) = flexAlignCenterCenter
            '''      .ColAlignment(5) = flexAlignRightCenter
            '''      .ColAlignment(6) = flexAlignCenterCenter
            '''      .ColAlignment(0) = flexAlignRightCenter
            '''
            '''      .FontFixed = "Small Fonts"
            '''      .Font = "MS Sans Serif"
            '''      .Font = 8.25
            '''      .FontFixed = 6.75
            '''
            '''      .Row = 0
            '''      .Col = 0
            '''      .text = "FOLIO"
            '''      .Col = 1
            '''      .text = "FECHA"
            '''      .Col = 2
            '''      .text = "CLIENTE"
            '''      .Col = 3
            '''      .text = "IMPORTE"
            '''      .Col = 4
            '''      .text = "ANTICIPO"
            '''      .Col = 5
            '''      .text = "ESTATUS"

            For I = 1 To 6
                .set_ColAlignmentFixed(I, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
            Next

            .Col = 1
            .ColSel = 6
        End With

    End Sub

    Sub Buscar()
        On Error GoTo Merr
        cWHERE = ""

        If chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            cWHERE = " And M.CodSucursal = " & intCodSucursal & "  "
        End If
        Select Case strTag
            Case "FRMABONODOCUMENTOS.TXTFOLIOVENTA"
                strCaptionForm = "Consulta de Folios de venta"
                strSQL = "   SELECT '' as x,  M.FolioVenta AS FOLIOVENTA, LTRIM(RTRIM(dbo.FormatFecha(M.FechaVenta, 5))) AS FECHA, C.DescCliente AS CLIENTE, " & "CASE M.Condicion WHEN 'CO' THEN 'Contado' WHEN 'CR' THEN 'Crédito' END AS TIPOVENTA, dbo.FormatCantidad(M.Total) AS TOTAL, " & "CASE M.Estatus WHEN 'C' THEN 'Cancelado' WHEN 'V' THEN 'Vigente' END AS ESTATUS " & "FROM         MovimientosVentasCab M, CatClientes C " & "WHERE     M.CodCliente = C.CodCliente AND TipoMovto =  'V' AND Condicion = 'CR' And DescCliente  LIKE '%" & Trim(txtNombre.Text) & "%' " & cWHERE & "ORDER BY M.FechaVenta DESC, m.folioventa DESC"

            Case "FRMABONOAPARTADOS.TXTFOLIO"
                strCaptionForm = "Consulta de Folios de apartado"
                strSQL = "select '' AS X, M.FolioVenta as FOLIO, dbo.FormatFecha( M.FechaVenta, 5) as FECHA, C.DescCliente as CLIENTE, " & "dbo.formatCantidad(M.Total + M.Redondeo) as TOTAL, dbo.FormatCantidad(M.Anticipo) as ANTICIPO, " & "CASE M.Estatus WHEN 'C' THEN 'Cancelado' WHEN 'V' THEN 'Vigente' END AS ESTATUS  " & "From MovimientosVentasCab M, CatClientes C Where M.CodCliente = C.CodCliente   And DescCliente  LIKE '%" & Trim(txtNombre.Text) & "%' " & cWHERE & "And TipoMovto= 'A' " & "Order by M.FechaVenta desc , m.folioventa Desc "

            Case "FRMDEVOLUCIONMERCANCIA.TXTFOLIOVENTA"
                strCaptionForm = "Consulta de folios de venta"
                strSQL = "Select '' AS X,  FolioVenta, Fecha, Cliente, TipoVenta,Total, Estatus from  dbo.VentasParaDevolucion(" & intCodSucursal & " ) " & "Where ((convert(money, TotalIngresos) >= convert(money,Total) AND (CONVERT(money, TOTAL) > 0))  " & "or  tipoventa='Contado' or  tipoventa='Crédito') And Cliente  LIKE '%" & Trim(txtNombre.Text) & "%' " & "Order by  FechaVenta Desc, FOLIOVENTA DESC"
                '''15SEP2006 - MAVF
            Case "FRMCONTROLREPARACIONES.TXTFOLIO"
                strCaptionForm = "Consulta de folios de Reparaciones"
                strSQL = "select '' AS X,  M.Folioreparacion as FOLIO, Ltrim(Rtrim(dbo.FormatFecha( M.FechaReparacion, 5))) as FECHA, C.DescCliente as CLIENTE, M.ImporteVta as IMPORTE, M.Anticipo as ANTICIPO, DBO.ReparacionesEstatus( M.CodSucursal, M.FechaReparacion, M.Folioreparacion ) AS ESTATUS " & "From Reparaciones M, CatClientes C Where M.CodCliente = C.CodCliente  And DescCliente  LIKE '%" & Trim(txtNombre.Text) & "%' " & cWHERE & "Order by M.FechaReparacion desc , M.folioReparacion Desc "
                '''15SEP2006 - MAVF
            Case "FRMCORPOCONTROLREPARACIONES_CORPO.TXTFOLIO"
                strCaptionForm = "Consulta de Folios de Reparaciones"
                strSQL = "select '' AS X , M.Folioreparacion as FOLIO, Ltrim(Rtrim(dbo.FormatFecha( M.FechaReparacion, 5))) as FECHA, C.DescCliente as CLIENTE, M.ImporteVta as IMPORTE, M.Anticipo as ANTICIPO, DBO.ReparacionesEstatus( M.CodSucursal, M.FechaReparacion, M.Folioreparacion ) AS ESTATUS " & "From Reparaciones M, CatClientes C Where M.CodSucursal <> 0 And M.CodCliente = C.CodCliente And DescCliente  LIKE '%" & Trim(txtNombre.Text) & "%' " & cWHERE & "Order by M.FechaReparacion desc , M.folioReparacion Desc "
            Case Else
                Exit Sub
        End Select

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, strSQL))
        RsGral = Cmd.Execute

        If RsGral.RecordCount > 0 Then
            Flexdet.Recordset = RsGral
        Else
            Flexdet.Rows = Flexdet.Rows + 1
            BorraFilas()
        End If
        If Flexdet.Rows <= 9 Then Flexdet.Rows = 9
        Me.Text = strCaptionForm

        With Me.Flexdet
            If strTag = "FRMCONTROLREPARACIONES.TXTFOLIO" Or strTag = "FRMCORPOCONTROLREPARACIONES_CORPO.TXTFOLIO" Then
                .set_ColWidth(0, 0, 0)
                .set_ColWidth(1, 0, 2000) 'fOLIO
                .set_ColWidth(2, 0, 1200) 'fECHA
                .set_ColWidth(3, 0, 3800) 'CLIENTE
                .set_ColWidth(4, 0, 1000) 'TIPO VENTA
                .set_ColWidth(5, 0, 1000) 'TOTAL
                .set_ColWidth(6, 0, 2000) 'ESTATUS
            Else
                .set_ColWidth(0, 0, 0)
                .set_ColWidth(1, 0, 2000) 'fOLIO
                .set_ColWidth(2, 0, 2000) 'fECHA
                .set_ColWidth(3, 0, 4000) 'CLIENTE
                .set_ColWidth(4, 0, 1000) 'TIPO VENTA
                .set_ColWidth(5, 0, 1000) 'TOTAL
                .set_ColWidth(6, 0, 1000) 'ESTATUS
            End If
            .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
            .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            .set_ColAlignment(4, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
            .set_ColAlignment(5, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColAlignment(6, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
            .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            For I = 1 To 6
                .set_ColAlignmentFixed(I, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
            Next
        End With
        Me.Flexdet.FontFixed = VB6.FontChangeName(Me.Flexdet.FontFixed, "Small Fonts")
        Me.Flexdet.Font = VB6.FontChangeName(Me.Flexdet.Font, "MS Sans Serif")
        Me.Flexdet.Font = VB6.FontChangeName(Me.Flexdet.Font, CStr(8.25))
        Me.Flexdet.FontFixed = VB6.FontChangeName(Me.Flexdet.FontFixed, CStr(6.75))
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub txtNombre_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.TextChanged
        If FueraChange Then Exit Sub
        If Trim(txtNombre.Text) = "" Then
            Encabezado()
            Buscar()
        End If
    End Sub

    Private Sub txtNombre_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNombre.Enter
        SelTextoTxt(txtNombre)
    End Sub

    Sub LlenaDatosSucursal()
        gStrSql = "SELECT      Ltrim(Rtrim(DescAlmacen)) as DescAlmacen From dbo.CatAlmacen Where CodAlmacen =" & intCodSucursal & "  And TipoAlmacen = 'P'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            FueraChange = True
            dbcSucursales.Text = RsGral.Fields("DescAlmacen").Value
            FueraChange = False
        End If
    End Sub

    Private Sub txtNombre_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtNombre.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Buscar()
            txtNombre.Focus()
        End If
    End Sub

    Sub BorraFilas()
        With Flexdet
            For I = 1 To .Rows - 2
                .RemoveItem(1)
            Next
        End With
    End Sub


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmConsultasEspeciales))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtNombre = New System.Windows.Forms.TextBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.chkTodasSucursales = New System.Windows.Forms.CheckBox()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me._Label1_3 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Flexdet = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Flexdet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtNombre
        '
        Me.txtNombre.AcceptsReturn = True
        Me.txtNombre.BackColor = System.Drawing.SystemColors.Window
        Me.txtNombre.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombre.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombre.Location = New System.Drawing.Point(176, 88)
        Me.txtNombre.MaxLength = 40
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombre.Size = New System.Drawing.Size(321, 20)
        Me.txtNombre.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtNombre, "Nombre")
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.Color.Black
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(8, 40)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(60, 17)
        Me._Label1_1.TabIndex = 1
        Me._Label1_1.Text = "Sucursal :"
        Me.ToolTip1.SetToolTip(Me._Label1_1, "Nombre de la Farmacia Actual")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtNombre)
        Me.Frame2.Controls.Add(Me.Frame1)
        Me.Frame2.Controls.Add(Me._Label1_3)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(8, 0)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(749, 121)
        Me.Frame2.TabIndex = 6
        Me.Frame2.TabStop = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.chkTodasSucursales)
        Me.Frame1.Controls.Add(Me.dbcSucursales)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(112, 8)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(409, 73)
        Me.Frame1.TabIndex = 7
        '
        'chkTodasSucursales
        '
        Me.chkTodasSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodasSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodasSucursales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodasSucursales.Location = New System.Drawing.Point(8, 16)
        Me.chkTodasSucursales.Name = "chkTodasSucursales"
        Me.chkTodasSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodasSucursales.Size = New System.Drawing.Size(155, 18)
        Me.chkTodasSucursales.TabIndex = 0
        Me.chkTodasSucursales.Text = "Todas las Sucursales"
        Me.chkTodasSucursales.UseVisualStyleBackColor = False
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(64, 40)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(323, 21)
        Me.dbcSucursales.TabIndex = 2
        '
        '_Label1_3
        '
        Me._Label1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_3, CType(3, Short))
        Me._Label1_3.Location = New System.Drawing.Point(120, 88)
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.Size = New System.Drawing.Size(50, 13)
        Me._Label1_3.TabIndex = 3
        Me._Label1_3.Text = "Cliente"
        '
        'Flexdet
        '
        Me.Flexdet.DataSource = Nothing
        Me.Flexdet.Location = New System.Drawing.Point(8, 127)
        Me.Flexdet.Name = "Flexdet"
        Me.Flexdet.OcxState = CType(resources.GetObject("Flexdet.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Flexdet.Size = New System.Drawing.Size(749, 166)
        Me.Flexdet.TabIndex = 8
        '
        'FrmConsultasEspeciales
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(768, 302)
        Me.Controls.Add(Me.Flexdet)
        Me.Controls.Add(Me.Frame2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(196, 148)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmConsultasEspeciales"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = " "
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Flexdet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private Sub dbcSucursales_SelectedValueChanged(sender As Object, e As EventArgs) Handles dbcSucursales.SelectedValueChanged
        PonerCodigoSucursal()
        Buscar()
    End Sub
End Class