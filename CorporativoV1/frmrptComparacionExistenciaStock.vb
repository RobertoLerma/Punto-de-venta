Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmrptComparacionExistenciaStock
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkSobrante As System.Windows.Forms.CheckBox
    Public WithEvents chkFaltante As System.Windows.Forms.CheckBox
    Public WithEvents chkIncluirSinDiferencia As System.Windows.Forms.CheckBox
    Public WithEvents fraImprimir As System.Windows.Forms.GroupBox
    Public WithEvents txtCodSucursal As System.Windows.Forms.TextBox
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaReporte As System.Windows.Forms.DateTimePicker
    Public WithEvents Frame5 As System.Windows.Forms.Panel
    Public WithEvents _lblVentas_5 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Dim mblnSalir As Boolean
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Dim mblnNuevo As Integer
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Dim mblnFueraChange As Boolean

    Sub Imprime()
        Dim rptComparacionExistenciaStock As New rptComparacionExistenciaStock
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Merr
        Dim aParam(2) As Object
        Dim aValues(2) As Object
        Dim cHAVING As String
        Dim Encabezado As String
        Dim lFaltSobr As String

        If ValidaDatos() = False Then Exit Sub
        Encabezado = "Comparación de Existencia y Stock Básico de Artículos"
        lFaltSobr = ""

        Select Case True
            '''arreglar el having
            Case chkFaltante.CheckState = System.Windows.Forms.CheckState.Checked And chkSobrante.CheckState = System.Windows.Forms.CheckState.Checked And chkIncluirSinDiferencia.CheckState = System.Windows.Forms.CheckState.Unchecked
                lFaltSobr = "X"
            Case chkFaltante.CheckState = System.Windows.Forms.CheckState.Checked And chkSobrante.CheckState = System.Windows.Forms.CheckState.Unchecked And chkIncluirSinDiferencia.CheckState = System.Windows.Forms.CheckState.Checked
                lFaltSobr = "F"
            Case chkFaltante.CheckState = System.Windows.Forms.CheckState.Checked And chkSobrante.CheckState = System.Windows.Forms.CheckState.Unchecked And chkIncluirSinDiferencia.CheckState = System.Windows.Forms.CheckState.Unchecked
                lFaltSobr = "F"
            Case chkFaltante.CheckState = System.Windows.Forms.CheckState.Unchecked And chkSobrante.CheckState = System.Windows.Forms.CheckState.Checked And chkIncluirSinDiferencia.CheckState = System.Windows.Forms.CheckState.Checked
                lFaltSobr = "S"
            Case chkFaltante.CheckState = System.Windows.Forms.CheckState.Unchecked And chkSobrante.CheckState = System.Windows.Forms.CheckState.Checked And chkIncluirSinDiferencia.CheckState = System.Windows.Forms.CheckState.Unchecked
                lFaltSobr = "S"
        End Select

        'Anterior
        'gStrSql = "SELECT     I.CodAlmacen, I.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " &
        '"AS Existencia, sk.Stock, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) - sk.Stock AS Dif, " &
        '"LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, Al.DescAlmacen, O.DescAlmacenOrigen, A.DescArticulo, G.DescGrupo " &
        '"FROM         dbo.Inventario I INNER JOIN " &
        '"dbo.CatArticulos A ON I.CodArticulo = A.CodArticulo INNER JOIN " &
        '"dbo.CatOrigen O ON A.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN " &
        '"dbo.CatAlmacen Al ON I.CodAlmacen = Al.CodAlmacen LEFT OUTER JOIN " &
        '"dbo.CatGrupos G ON A.CodGrupo = G.CodGrupo FULL OUTER JOIN " &
        '"dbo.StockPorArticulo() sk ON A.CodArticulo = sk.CodArticulo CROSS JOIN " &
        '"dbo.ConfiguracionGeneral " &
        '"GROUP BY I.CodArticulo, A.CodArticulo, A.DescArticulo, A.CodGrupo, A.CodUnidad, O.DescAlmacenOrigen, Al.DescAlmacen, I.CodAlmacen, I.CodAlmacenOrigen, " &
        '"A.CostoReal , dbo.ConfiguracionGeneral.NombreEmp, A.DescArticulo, G.DescGrupo, sk.Stock " &
        '"Having (I.CodAlmacen = " & Numerico(txtCodSucursal) & ") " & cHAVING &
        '"ORDER BY I.CodArticulo"

        gStrSql = "(SELECT I.CodAlmacen, I.CodAlmacenOrigen, I.CodArticulo, SK.CodGrupo, sum((I.ExistenciaInicial + I.Entradas) - (I.Salidas + I.Apartados)) AS Existencia, SK.Stock, Abs((SUM((I.ExistenciaInicial + I.Entradas) - (I.Salidas + I.Apartados)) - sk.Stock)) as Dif, '" & gstrCorpoNOMBREEMPRESA & "' as NombreEmp, Al.DescAlmacen, O.DescAlmacenOrigen, A.DescArticulo, G.DescGrupo, '' as Nivel1, Case When ltrim(rtrim(F.DescFamilia)) = '' Then ltrim(rtrim(L.DescLinea)) Else ltrim(rtrim(F.DescFamilia)) + ' - ' + ltrim(rtrim(L.DescLinea)) End as Nivel2 " & "FROM    dbo.Inventario I (Nolock) Left OUTER JOIN dbo.StockPorArticulo_Grupo (1) SK  ON I.CodArticulo = SK.CodArticulo INNER JOIN dbo.CatArticulos A (Nolock) ON I.CodArticulo = A.CodArticulo INNER JOIN dbo.CatOrigen O (Nolock) ON I.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN dbo.CatAlmacen Al (Nolock) ON I.CodAlmacen = Al.CodAlmacen LEFT OUTER JOIN dbo.CatGrupos G (Nolock) ON A.CodGrupo = G.CodGrupo Inner Join CatFamilias F On SK.CodGrupo = F.CodGrupo And SK.CodFamilia = F.CodFamilia Inner Join CatLineas L On SK.CodGrupo = L.CodGrupo And SK.CodFamilia = L.CodFamilia And SK.CodLinea = L.CodLinea " & "Where   Al.TipoAlmacen = 'P' " & "GROUP   BY I.CodAlmacen, I.CodAlmacenOrigen, I.CodArticulo, SK.CodGrupo, SK.CodFamilia, SK.CodLinea, sk.Stock , Al.DescAlmacen, O.DescAlmacenOrigen, A.DescArticulo, G.DescGrupo, F.DescFamilia, L.DescLinea " & "Having  (I.CodAlmacen = " & CInt(Numerico(txtCodSucursal.Text)) & ")  ) "
        gStrSql = gStrSql & " UNION " & "(SELECT I.CodAlmacen, I.CodAlmacenOrigen, I.CodArticulo, SK.CodGrupo, sum((I.ExistenciaInicial + I.Entradas) - (I.Salidas + I.Apartados)) AS Existencia, SK.Stock, Abs((SUM((I.ExistenciaInicial + I.Entradas) - (I.Salidas + I.Apartados)) - sk.Stock)) as Dif, '" & gstrCorpoNOMBREEMPRESA & "' as NombreEmp, Al.DescAlmacen, O.DescAlmacenOrigen, A.DescArticulo, G.DescGrupo, '' as Nivel1, Case When ltrim(rtrim(F.DescMarca)) = '' Then '' Else ltrim(rtrim(F.DescMarca)) End as Nivel2 " & "FROM    dbo.Inventario I (Nolock) Left OUTER JOIN dbo.StockPorArticulo_Grupo (2) SK  ON I.CodArticulo = SK.CodArticulo INNER JOIN dbo.CatArticulos A (Nolock) ON I.CodArticulo = A.CodArticulo INNER JOIN dbo.CatOrigen O (Nolock) ON I.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN dbo.CatAlmacen Al (Nolock) ON I.CodAlmacen = Al.CodAlmacen LEFT OUTER JOIN dbo.CatGrupos G (Nolock) ON A.CodGrupo = G.CodGrupo Inner Join CatMarcas F On SK.CodGrupo = F.CodGrupo And SK.CodMarca = F.CodMarca " & "Where   Al.TipoAlmacen = 'P' " & "GROUP   BY I.CodAlmacen, I.CodAlmacenOrigen, I.CodArticulo, SK.CodGrupo, SK.CodMarca, SK.CodModelo, sk.Stock , Al.DescAlmacen, O.DescAlmacenOrigen, A.DescArticulo, G.DescGrupo, F.DescMarca " & "Having  (I.CodAlmacen = " & CInt(Numerico(txtCodSucursal.Text)) & ")  ) "
        gStrSql = gStrSql & " UNION " & "(SELECT I.CodAlmacen, I.CodAlmacenOrigen, I.CodArticulo, SK.CodGrupo, sum((I.ExistenciaInicial + I.Entradas) - (I.Salidas + I.Apartados)) AS Existencia, SK.Stock, Abs((SUM((I.ExistenciaInicial + I.Entradas) - (I.Salidas + I.Apartados)) - sk.Stock)) as Dif, '" & gstrCorpoNOMBREEMPRESA & "' as NombreEmp, Al.DescAlmacen, O.DescAlmacenOrigen, A.DescArticulo, G.DescGrupo, '' as Nivel1, Case When ltrim(rtrim(F.DescFamilia)) = '' Then '' Else ltrim(rtrim(F.DescFamilia)) End as Nivel2 " & "FROM    dbo.Inventario I (Nolock) Left OUTER JOIN dbo.StockPorArticulo_Grupo (3) SK  ON I.CodArticulo = SK.CodArticulo INNER JOIN dbo.CatArticulos A (Nolock) ON I.CodArticulo = A.CodArticulo INNER JOIN dbo.CatOrigen O (Nolock) ON I.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN dbo.CatAlmacen Al (Nolock) ON I.CodAlmacen = Al.CodAlmacen LEFT OUTER JOIN dbo.CatGrupos G (Nolock) ON A.CodGrupo = G.CodGrupo Inner Join CatFamilias F On SK.CodGrupo = F.CodGrupo And SK.CodFamilia = F.CodFamilia " & "Where   Al.TipoAlmacen = 'P' " & "GROUP   BY I.CodAlmacen, I.CodAlmacenOrigen, I.CodArticulo, SK.CodGrupo, SK.CodFamilia, SK.CodLinea, sk.Stock , Al.DescAlmacen, O.DescAlmacenOrigen, A.DescArticulo, G.DescGrupo, F.DescFamilia " & "Having  (I.CodAlmacen = " & CInt(Numerico(txtCodSucursal.Text)) & ")  ) "

        gStrSql = gStrSql & "ORDER   BY sk.stock"

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
            rptComparacionExistenciaStock.SetDataSource(frmReportes.rsReport)
        End If

        'aParam(1) = "EncabezadoReporte"
        'aValues(1) = Encabezado
        'aParam(2) = "FaltSobr"
        'aValues(2) = lFaltSobr
        'frmReportes.Report = rptComparacionExistenciaStock 'Es el nombre del archivo que se incluyó en el proyecto
        'frmReportes.Imprime("", aParam, aValues)

        If (Encabezado <> Nothing) Then
            pdvNum.Value = Encabezado : pvNum.Add(pdvNum)
            rptComparacionExistenciaStock.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptComparacionExistenciaStock.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
        End If

        If (lFaltSobr <> Nothing) Then
            pdvNum.Value = lFaltSobr : pvNum.Add(pdvNum)
            rptComparacionExistenciaStock.DataDefinition.ParameterFields("FaltSobr").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptComparacionExistenciaStock.DataDefinition.ParameterFields("FaltSobr").ApplyCurrentValues(pvNum)
        End If

        frmReportes.reporteActual = rptComparacionExistenciaStock
        frmReportes.Show()


Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcSucursales_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursales" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
        mblnFueraChange = True
        txtCodSucursal.Text = ""
        mblnFueraChange = False
    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursales.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen where TipoAlmacen ='P'ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
    End Sub

    Private Sub dbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        ElseIf eventArgs.KeyCode = System.Windows.Forms.Keys.Return Then
            AvanzarTab(Me)
        End If
    End Sub


    Private Sub dbcSucursales_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyUp
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Up Or eventArgs.KeyCode = System.Windows.Forms.Keys.Down Then
            PonerCodigoSucursal()
            Exit Sub
        End If
    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        mblnFueraChange = True
        If intCodSucursal = 0 Then
            txtCodSucursal.Text = ""
        Else
            txtCodSucursal.Text = VB6.Format(intCodSucursal, "000")
        End If
        mblnFueraChange = False
        If EsAlmacenGral(CInt(Numerico(txtCodSucursal.Text))) Then
            MsgBox("Ha seleccionado el almacén general del corporativo. Es probable que la información obtenida en el reporte no sea congruente." & vbNewLine & "Debido a que en el almacén general se concentra toda la información de las sucursales.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            '        dbcSucursales.SetFocus
        End If
    End Sub

    Function ValidaDatos() As Boolean
        If Trim(dbcSucursales.Text) = "" Then
            MsgBox("Proporcione la sucursal sobre la cual se buscará la información.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            dbcSucursales.Focus()
            Exit Function
        End If
        If chkFaltante.CheckState = System.Windows.Forms.CheckState.Unchecked And chkSobrante.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            MsgBox("Seleccione una opción para la diferencia.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            chkFaltante.Focus()
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Private Sub dtpFechaIncio_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            KeyCode = 0
        End If
    End Sub

    Private Sub dbcSucursales_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcSucursales.MouseUp
        PonerCodigoSucursal()
    End Sub

    Private Sub frmrptComparacionExistenciaStock_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmrptComparacionExistenciaStock_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub Form_Initialize_Renamed()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmrptComparacionExistenciaStock_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        ' En este evento del formulario se valida la tecla presionada.
        ' Si es Enter se simula un tab(Avanza al siguiente control)
        ' Si es Escape, se simula un Retroceso de TAB (Regresa al control anterior)
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmrptComparacionExistenciaStock_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub Nuevo()
        mblnFueraChange = True
        txtCodSucursal.Text = ""
        dbcSucursales.Text = ""
        dtpFechaReporte.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
        mblnFueraChange = False
        chkFaltante.CheckState = System.Windows.Forms.CheckState.Checked
        chkSobrante.CheckState = System.Windows.Forms.CheckState.Checked
        chkIncluirSinDiferencia.CheckState = System.Windows.Forms.CheckState.Checked
    End Sub

    Private Sub frmrptComparacionExistenciaStock_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Nuevo()
        '    LlenaDatosSucursal (gintCodAlmacen)
    End Sub

    Private Sub frmrptComparacionExistenciaStock_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSalir Then
        '    'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        '    ModEstandar.RestaurarForma(Me, False)
        '    Cancel = 0 'Para que no salga del Formulario hasta que guarde los datos, si no tiene premiso de hacerlo
        'Else 'Se quiere salir con escape
        '    mblnSalir = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
        '        Case MsgBoxResult.Yes 'Sale del Formulario, pero antes preguntar si desea grabar los datos registrados, solo cuando es nuevo
        '            Cancel = 0 'Sale de la Captura, Con 1: Sigue en la captura
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmrptComparacionExistenciaStock_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Sub Limpiar()
        Nuevo()
        txtCodSucursal.Focus()
    End Sub

    Private Sub txtCodSucursal_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.TextChanged
        If mblnFueraChange = True Then Exit Sub
        mblnFueraChange = True
        dbcSucursales.Text = ""
        mblnFueraChange = False
    End Sub

    Private Sub txtCodSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Enter
        SelTextoTxt(txtCodSucursal)
    End Sub

    Private Sub txtCodSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodSucursal.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            Exit Sub
        End If
    End Sub

    Private Sub txtCodsucursal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodsucursal.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Leave
        LlenaDatosSucursal()
        If EsAlmacenGral(CInt(Numerico(txtCodSucursal.Text))) Then
            MsgBox("Ha seleccionado el almacén general del corporativo. Es probable que la información obtenida en el reporte no sea congruente." & vbNewLine & "Debido a que en el almacén general se concentra toda la información de las sucursales.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
        End If
    End Sub

    Sub LlenaDatosSucursal()
        On Error GoTo Merr
        If CDbl(Numerico(Trim(txtCodSucursal.Text))) = 0 Then Exit Sub
        gStrSql = "SELECT      Ltrim(Rtrim(DescAlmacen)) as DescAlmacen From dbo.CatAlmacen Where CodAlmacen =" & Numerico(txtCodSucursal.Text) & "  And TipoAlmacen = 'P'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            mblnFueraChange = True
            dbcSucursales.Text = RsGral.Fields("DescAlmacen").Value
            mblnFueraChange = False
        Else
            MsgBox("Código de sucursal no existe." & vbNewLine & "Verifique Por Favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            txtCodSucursal.Text = ""
            txtCodSucursal.Focus()
            Exit Sub
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function EsAlmacenGral(ByRef CodAlmacen As Integer) As Boolean
        On Error GoTo Merr
        If CDbl(Numerico(Trim(txtCodSucursal.Text))) = 0 Then Exit Function
        EsAlmacenGral = False
        gStrSql = "SELECT   AlmGral  From dbo.CatAlmacen Where CodAlmacen =" & Numerico(txtCodSucursal.Text) & "  And TipoAlmacen = 'P'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            If RsGral.Fields("AlmGral").Value = True Then
                EsAlmacenGral = True
            End If
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Sub PonerCodigoSucursal()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'DCLostFocus dbcAlmacenSalida, gStrSql, intCodSucursal
        mblnFueraChange = True
        If RsGral.RecordCount <= 0 Then
            txtCodSucursal.Text = ""
        Else
            txtCodSucursal.Text = VB6.Format(RsGral.Fields("CodAlmacen").Value, "000")
        End If
        mblnFueraChange = False
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.fraImprimir = New System.Windows.Forms.GroupBox()
        Me.chkSobrante = New System.Windows.Forms.CheckBox()
        Me.chkFaltante = New System.Windows.Forms.CheckBox()
        Me.chkIncluirSinDiferencia = New System.Windows.Forms.CheckBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtCodSucursal = New System.Windows.Forms.TextBox()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me.Frame5 = New System.Windows.Forms.Panel()
        Me.dtpFechaReporte = New System.Windows.Forms.DateTimePicker()
        Me._lblVentas_5 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        Me.fraImprimir.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.Frame5.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.Color.Black
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(8, 18)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(60, 17)
        Me._Label1_1.TabIndex = 2
        Me._Label1_1.Text = "Sucursal :"
        Me.ToolTip1.SetToolTip(Me._Label1_1, "Nombre de la Farmacia Actual")
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.fraImprimir)
        Me.Frame2.Controls.Add(Me.Frame1)
        Me.Frame2.Controls.Add(Me.Frame5)
        Me.Frame2.Controls.Add(Me._lblVentas_5)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(10, 4)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(409, 210)
        Me.Frame2.TabIndex = 8
        Me.Frame2.TabStop = False
        '
        'fraImprimir
        '
        Me.fraImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.fraImprimir.Controls.Add(Me.chkSobrante)
        Me.fraImprimir.Controls.Add(Me.chkFaltante)
        Me.fraImprimir.Controls.Add(Me.chkIncluirSinDiferencia)
        Me.fraImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraImprimir.Location = New System.Drawing.Point(184, 96)
        Me.fraImprimir.Name = "fraImprimir"
        Me.fraImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraImprimir.Size = New System.Drawing.Size(217, 101)
        Me.fraImprimir.TabIndex = 5
        Me.fraImprimir.TabStop = False
        Me.fraImprimir.Text = " Diferencia "
        '
        'chkSobrante
        '
        Me.chkSobrante.BackColor = System.Drawing.SystemColors.Control
        Me.chkSobrante.Checked = True
        Me.chkSobrante.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSobrante.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSobrante.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSobrante.Location = New System.Drawing.Point(96, 36)
        Me.chkSobrante.Name = "chkSobrante"
        Me.chkSobrante.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSobrante.Size = New System.Drawing.Size(77, 23)
        Me.chkSobrante.TabIndex = 7
        Me.chkSobrante.Text = "Sobrante"
        Me.chkSobrante.UseVisualStyleBackColor = False
        '
        'chkFaltante
        '
        Me.chkFaltante.BackColor = System.Drawing.SystemColors.Control
        Me.chkFaltante.Checked = True
        Me.chkFaltante.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFaltante.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkFaltante.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkFaltante.Location = New System.Drawing.Point(96, 16)
        Me.chkFaltante.Name = "chkFaltante"
        Me.chkFaltante.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkFaltante.Size = New System.Drawing.Size(77, 22)
        Me.chkFaltante.TabIndex = 6
        Me.chkFaltante.Text = "Faltante"
        Me.chkFaltante.UseVisualStyleBackColor = False
        '
        'chkIncluirSinDiferencia
        '
        Me.chkIncluirSinDiferencia.BackColor = System.Drawing.SystemColors.Control
        Me.chkIncluirSinDiferencia.Checked = True
        Me.chkIncluirSinDiferencia.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkIncluirSinDiferencia.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIncluirSinDiferencia.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkIncluirSinDiferencia.Location = New System.Drawing.Point(96, 56)
        Me.chkIncluirSinDiferencia.Name = "chkIncluirSinDiferencia"
        Me.chkIncluirSinDiferencia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIncluirSinDiferencia.Size = New System.Drawing.Size(77, 33)
        Me.chkIncluirSinDiferencia.TabIndex = 11
        Me.chkIncluirSinDiferencia.Text = "Incluir sin diferencia"
        Me.chkIncluirSinDiferencia.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtCodSucursal)
        Me.Frame1.Controls.Add(Me.dbcSucursales)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 40)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(393, 49)
        Me.Frame1.TabIndex = 10
        Me.Frame1.TabStop = False
        '
        'txtCodSucursal
        '
        Me.txtCodSucursal.AcceptsReturn = True
        Me.txtCodSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucursal.Location = New System.Drawing.Point(64, 16)
        Me.txtCodSucursal.MaxLength = 5
        Me.txtCodSucursal.Name = "txtCodSucursal"
        Me.txtCodSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucursal.Size = New System.Drawing.Size(49, 21)
        Me.txtCodSucursal.TabIndex = 3
        Me.txtCodSucursal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(120, 16)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(267, 21)
        Me.dbcSucursales.TabIndex = 4
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.dtpFechaReporte)
        Me.Frame5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame5.Enabled = False
        Me.Frame5.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.Frame5.Location = New System.Drawing.Point(288, 8)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(111, 42)
        Me.Frame5.TabIndex = 9
        '
        'dtpFechaReporte
        '
        Me.dtpFechaReporte.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaReporte.Location = New System.Drawing.Point(12, 8)
        Me.dtpFechaReporte.Name = "dtpFechaReporte"
        Me.dtpFechaReporte.Size = New System.Drawing.Size(95, 20)
        Me.dtpFechaReporte.TabIndex = 1
        '
        '_lblVentas_5
        '
        Me._lblVentas_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblVentas.SetIndex(Me._lblVentas_5, CType(5, Short))
        Me._lblVentas_5.Location = New System.Drawing.Point(240, 20)
        Me._lblVentas_5.Name = "_lblVentas_5"
        Me._lblVentas_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_5.Size = New System.Drawing.Size(49, 12)
        Me._lblVentas_5.TabIndex = 0
        Me._lblVentas_5.Text = "Fecha :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(125, 230)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 99
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(10, 230)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 98
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmrptComparacionExistenciaStock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(427, 279)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(368, 237)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmrptComparacionExistenciaStock"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Comparación de Existencias y Stock"
        Me.Frame2.ResumeLayout(False)
        Me.fraImprimir.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.Frame5.ResumeLayout(False)
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class