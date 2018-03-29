'**********************************************************************************************************************'
'*PROGRAMA: ABC DE ARTICULOS JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmCorpoABCArticulos
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkCodigoAnterior As System.Windows.Forms.CheckBox
    Public WithEvents txtCodArtAnterior As System.Windows.Forms.TextBox
    Public WithEvents dbcOrigen As System.Windows.Forms.ComboBox
    Public WithEvents _lblArticulo_32 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_31 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents txtDescArticulo As System.Windows.Forms.TextBox
    Public WithEvents txtCodArticulo As System.Windows.Forms.TextBox
    Public WithEvents txtMDSCertificado As System.Windows.Forms.TextBox
    Public WithEvents txtMDSPureza As System.Windows.Forms.TextBox
    Public WithEvents txtMDSColor As System.Windows.Forms.TextBox
    Public WithEvents txtMDSPeso As System.Windows.Forms.TextBox
    Public WithEvents lblEstatus As System.Windows.Forms.Label
    Public WithEvents lblMDSCertificado As System.Windows.Forms.Label
    Public WithEvents lblMDSPureza As System.Windows.Forms.Label
    Public WithEvents lblMDSColor As System.Windows.Forms.Label
    Public WithEvents lblMDSPeso As System.Windows.Forms.Label
    Public WithEvents fraDiamanteSuelto As System.Windows.Forms.GroupBox
    Public WithEvents _txtAdicional_0 As System.Windows.Forms.TextBox
    Public WithEvents _optMoneda_11 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_10 As System.Windows.Forms.RadioButton
    Public WithEvents _fraMoneda_5 As System.Windows.Forms.GroupBox
    Public WithEvents _optMoneda_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_0 As System.Windows.Forms.RadioButton
    Public WithEvents _fraMoneda_0 As System.Windows.Forms.GroupBox
    Public WithEvents Image1 As System.Windows.Forms.PictureBox
    Public WithEvents _fraImagen_0 As System.Windows.Forms.GroupBox
    Public WithEvents _cmdBuscarImagen_0 As System.Windows.Forms.Button
    Public WithEvents _txtImagen_0 As System.Windows.Forms.TextBox
    Public WithEvents _Frame4_0 As System.Windows.Forms.GroupBox
    Public WithEvents _txtCodigodelProveedor_0 As System.Windows.Forms.TextBox
    Public WithEvents _dbcProveedor_0 As System.Windows.Forms.ComboBox
    Public WithEvents _cboUnidad_0 As System.Windows.Forms.ComboBox
    Public WithEvents _cboAlmacen_0 As System.Windows.Forms.ComboBox
    Public WithEvents _lblArticulo_36 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_35 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_11 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_10 As System.Windows.Forms.Label
    Public WithEvents _Frame2_0 As System.Windows.Forms.GroupBox
    Public WithEvents _txtDescripcion_0 As System.Windows.Forms.TextBox
    Public WithEvents _txtCostoReal_0 As System.Windows.Forms.TextBox
    Public WithEvents _txtPrecioenDolares_0 As System.Windows.Forms.TextBox
    Public WithEvents _txtCostoIndirecto_0 As System.Windows.Forms.TextBox
    Public WithEvents _txtCostoAdicional_0 As System.Windows.Forms.TextBox
    Public WithEvents _txtCostoFactura_0 As System.Windows.Forms.TextBox
    Public WithEvents _lblMargen_0 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_34 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_5 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_8 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_7 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_6 As System.Windows.Forms.Label
    Public WithEvents _Frame1_0 As System.Windows.Forms.GroupBox
    Public WithEvents _dbcFamilia_0 As System.Windows.Forms.ComboBox
    Public WithEvents _dbcLinea_0 As System.Windows.Forms.ComboBox
    Public WithEvents dbcSubLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcKilates As System.Windows.Forms.ComboBox
    Public WithEvents _dbcMaterial_0 As System.Windows.Forms.ComboBox
    Public WithEvents _lblArticulo_33 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_9 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_29 As System.Windows.Forms.Label
    Public WithEvents _lblDescripcion_0 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_26 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_37 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_4 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_3 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_2 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_1 As System.Windows.Forms.Label
    Public WithEvents _fraContenedor_0 As System.Windows.Forms.Panel
    Public WithEvents _sstArticulo_TabPage0 As System.Windows.Forms.TabPage
    Public WithEvents _txtAdicional_1 As System.Windows.Forms.TextBox
    Public WithEvents _optMoneda_7 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_6 As System.Windows.Forms.RadioButton
    Public WithEvents _fraMoneda_3 As System.Windows.Forms.GroupBox
    Public WithEvents _txtDescripcion_1 As System.Windows.Forms.TextBox
    Public WithEvents _optGenero_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optGenero_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optGenero_2 As System.Windows.Forms.RadioButton
    Public WithEvents _fraArticulo_1 As System.Windows.Forms.GroupBox
    Public WithEvents _optMovimiento_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optMovimiento_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optMovimiento_2 As System.Windows.Forms.RadioButton
    Public WithEvents _fraArticulo_2 As System.Windows.Forms.GroupBox
    Public WithEvents Image2 As System.Windows.Forms.PictureBox
    Public WithEvents _fraImagen_1 As System.Windows.Forms.GroupBox
    Public WithEvents _optMoneda_3 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_2 As System.Windows.Forms.RadioButton
    Public WithEvents _fraMoneda_1 As System.Windows.Forms.GroupBox
    Public WithEvents _txtCostoReal_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtPrecioenDolares_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtCostoIndirecto_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtCostoAdicional_1 As System.Windows.Forms.TextBox
    Public WithEvents _txtCostoFactura_1 As System.Windows.Forms.TextBox
    Public WithEvents _lblMargen_1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_40 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_41 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_42 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_43 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_44 As System.Windows.Forms.Label
    Public WithEvents _Frame1_2 As System.Windows.Forms.GroupBox
    Public WithEvents _txtImagen_1 As System.Windows.Forms.TextBox
    Public WithEvents _cmdBuscarImagen_1 As System.Windows.Forms.Button
    Public WithEvents _Frame4_1 As System.Windows.Forms.GroupBox
    Public WithEvents _txtCodigodelProveedor_1 As System.Windows.Forms.TextBox
    Public WithEvents _dbcProveedor_1 As System.Windows.Forms.ComboBox
    Public WithEvents _cboUnidad_1 As System.Windows.Forms.ComboBox
    Public WithEvents _cboAlmacen_1 As System.Windows.Forms.ComboBox
    Public WithEvents _lblArticulo_18 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_19 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_20 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_21 As System.Windows.Forms.Label
    Public WithEvents _Frame2_2 As System.Windows.Forms.GroupBox
    Public WithEvents chkCrono As System.Windows.Forms.CheckBox
    Public WithEvents dbcMarca As System.Windows.Forms.ComboBox
    Public WithEvents dbcModelo As System.Windows.Forms.ComboBox
    Public WithEvents _dbcMaterial_1 As System.Windows.Forms.ComboBox
    Public WithEvents _lblArticulo_45 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_27 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_13 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_12 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_14 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_15 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_16 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_17 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_38 As System.Windows.Forms.Label
    Public WithEvents _lblDescripcion_1 As System.Windows.Forms.Label
    Public WithEvents _fraContenedor_1 As System.Windows.Forms.Panel
    Public WithEvents _sstArticulo_TabPage1 As System.Windows.Forms.TabPage
    Public WithEvents _txtAdicional_2 As System.Windows.Forms.TextBox
    Public WithEvents _optMoneda_9 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_8 As System.Windows.Forms.RadioButton
    Public WithEvents _fraMoneda_4 As System.Windows.Forms.GroupBox
    Public WithEvents _optMoneda_4 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_5 As System.Windows.Forms.RadioButton
    Public WithEvents _fraMoneda_2 As System.Windows.Forms.GroupBox
    Public WithEvents _txtImagen_2 As System.Windows.Forms.TextBox
    Public WithEvents _cmdBuscarImagen_2 As System.Windows.Forms.Button
    Public WithEvents _Frame4_2 As System.Windows.Forms.GroupBox
    Public WithEvents _txtCodigodelProveedor_2 As System.Windows.Forms.TextBox
    Public WithEvents _dbcProveedor_2 As System.Windows.Forms.ComboBox
    Public WithEvents _cboUnidad_2 As System.Windows.Forms.ComboBox
    Public WithEvents _cboAlmacen_2 As System.Windows.Forms.ComboBox
    Public WithEvents _lblArticulo_50 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_51 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_52 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_53 As System.Windows.Forms.Label
    Public WithEvents _Frame2_3 As System.Windows.Forms.GroupBox
    Public WithEvents _txtCostoFactura_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtCostoAdicional_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtCostoIndirecto_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtPrecioenDolares_2 As System.Windows.Forms.TextBox
    Public WithEvents _txtCostoReal_2 As System.Windows.Forms.TextBox
    Public WithEvents _lblMargen_2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_22 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_23 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_46 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_47 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_48 As System.Windows.Forms.Label
    Public WithEvents _Frame1_3 As System.Windows.Forms.GroupBox
    Public WithEvents _txtDescripcion_2 As System.Windows.Forms.TextBox
    Public WithEvents Image3 As System.Windows.Forms.PictureBox
    Public WithEvents _fraImagen_2 As System.Windows.Forms.GroupBox
    Public WithEvents _dbcFamilia_1 As System.Windows.Forms.ComboBox
    Public WithEvents _dbcLinea_1 As System.Windows.Forms.ComboBox
    Public WithEvents _dbcMaterial_2 As System.Windows.Forms.ComboBox
    Public WithEvents _lblArticulo_54 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_49 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_28 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_24 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_25 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_30 As System.Windows.Forms.Label
    Public WithEvents _lblArticulo_39 As System.Windows.Forms.Label
    Public WithEvents _lblDescripcion_2 As System.Windows.Forms.Label
    Public WithEvents _fraContenedor_2 As System.Windows.Forms.Panel
    Public WithEvents _sstArticulo_TabPage2 As System.Windows.Forms.TabPage
    Public WithEvents sstArticulo As System.Windows.Forms.TabControl
    Public WithEvents _lblArticulo_0 As System.Windows.Forms.Label
    Public WithEvents Frame1 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents Frame2 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents Frame4 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents cboAlmacen As System.Windows.Forms.ComboBox
    Public WithEvents cboUnidad As System.Windows.Forms.ComboBox
    Public WithEvents cmdBuscarImagen As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
    Public WithEvents dbcFamilia As System.Windows.Forms.ComboBox
    Public WithEvents dbcLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcMaterial As System.Windows.Forms.ComboBox
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents fraArticulo As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents fraContenedor As Microsoft.VisualBasic.Compatibility.VB6.PanelArray
    Public WithEvents fraImagen As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents fraMoneda As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblArticulo As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblDescripcion As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblMargen As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optGenero As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents optMovimiento As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents txtAdicional As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtCodigodelProveedor As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtCostoAdicional As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtCostoFactura As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtCostoIndirecto As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtCostoReal As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtDescripcion As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtImagen As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents txtPrecioenDolares As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    Public WithEvents btnLimpiar As Button
    Public WithEvents btnEliminar As Button
    Public WithEvents btnGuardar As Button
    Public strControlActual As String 'Nombre del control actual


    ''' ****************************************************************************************************************************************************
    ''' CORRECCION FUNCIONAMIENTO DE DESCRIPCION DEL ARTICULO PARA JOYERIA, YA QUE CUANDO SE SELECCIONABA UNA SUBLINEA EL
    ''' DATO CORRESPONDIENTE A LA DESCRIPCION CORTE NO SE ACTUALIZABA EN LA DESCRIPCION DEL ARTICULO.
    ''' 06AGO2007 - MAVF
    '''
    ''' SE ELIMINÓ VALIDACIÓN DE KILATES PARA GRUPO DE JOYERIA, ESTO PERMITE QUE EL DATO 0 KILATES NO SE CONSIDERE EN LA DESCRIPCION
    ''' 26MAR2008 - MAVF
    '''
    ''' SE AGREGARON 4 CAMPOS NUEVOS PARA EL MANEJO DE DIAMANTE SUELTO ( MDS )
    ''' SE MODIFICO FUNCIONES:  GUARDAR - VALIDARDATOSMANEJODIAMANTESUELTO-CAMBIOS-NUEVO-LLENADATOS-ELIMINAR
    ''' SPS --> ModStoredProcedures.PR_IMECatArticulos  -- OPCION Guarar-Modificar
    ''' 27OCT2010 - MAVF Ver
    '''
    ''' MDS CORRECCION - SE ELIMINO VALIDACIÓN DE PESO 0.00 YA QUE ES UN DATO NO REQUERIDO PARA TODOS LOS ARTICULOS DEL TIPO JOYERIA ( SOLO APLICA EN JOYERIA DIAMANTE SUELTO )
    ''' SE CONSIDERO ADICIONALMENTE  -SIN KILATAJE-  PARA EVITAR QUE PONGA 0K EN LA DESCRIPCIÓN - PETICION DE MRB DE ULTIMA HORA ( SIN $$$)
    ''' 08NOV2010 - MAVF Ver
    '''
    ''' Ver 1.1       Estatus:  Aprobado
    ''' ****************************************************************************************************************************************************


    Const cINDEFINIDA As String = "[ Vacío ... ]" 'Procure que no haya espacios en blanco en los extremos
    Const cINDEFINIDO As String = "[ Vacío ... ]" 'Procure que no haya espacios en blanco en los extremos
    Const cSINKILATES As String = "(SIN KILATES)" 'Para evitar poner Kiltaje en 0k    '''08NOV2010 - MAVF

    Const nJOYERIA As Integer = 0
    Const nRELOJERIA As Integer = 1
    Const nVARIOS As Integer = 2

    Const C_CRONO As String = "CHR"

    Dim mblnSalir As Boolean
    Dim ResBusquedaArt As Integer
    'Variables para controlar la Descripción del Artículo
    Dim cDescripcion As String
    Dim cFamilia(1) As String 'Tanto para Joyería como para Varios
    Dim cLinea(1) As String 'Tanto para Joyería como para Varios
    Dim cSubLinea As String
    Dim cSubLineaDescCorta As String
    Dim cKilates As String
    Dim cMarca As String
    Dim cModelo As String
    Dim cTipoMaterial As String
    Dim cTipoMaterialDescCorta As String
    Dim cMovimiento As String
    Dim cGenero As String
    Dim cCrono As String
    Dim lCrono As Boolean
    Dim cMovimientoTag As String
    Dim cGeneroTag As String
    Dim cCronoTag As String
    Dim cMonedaCompra As String
    Dim cMonedaCompraTag As String
    Dim cOrigen As String
    Dim cOrigenTag As String
    Dim cCodProveedor As String

    Dim nCostoFactura As Decimal
    Dim nCostoFacturaPesos As Decimal
    Dim nCostoFacturaTag As Decimal
    Dim nCostoAdicional As Decimal
    Dim nCostoAdicionalPesos As Decimal
    Dim nCostoAdicionalTag As Decimal
    Dim nCostoIndirectos As Decimal
    Dim nCostoIndirectosPesos As Decimal
    Dim nCostoIndirectosTag As Decimal
    Dim nCostoReal As Decimal
    Dim nCostoRealTag As Decimal

    Dim mblnCambiosEnCodigo As Boolean
    Dim mblnNuevo As Boolean
    Dim mblnLlenoDatos As Boolean

    'Variables para los combos de Joyería y Varios, y el tipo de Material
    Dim tecla As Integer
    Dim mintCodArticulo As Integer

    Dim mintJFam As Integer
    Dim mintJLin As Integer
    Dim mintJSub As Integer
    Dim mintVFam As Integer
    Dim mintVLin As Integer
    Dim mintRMar As Integer
    Dim mintRMod As Integer
    Dim mintJMaterial As Integer
    Dim mintRMaterial As Integer
    Dim mintVMaterial As Integer
    Dim mintJProv As Integer
    Dim mintRProv As Integer
    Dim mintVProv As Integer
    Dim mintJUnidad As Integer
    Dim mintRUnidad As Integer
    Dim mintVUnidad As Integer
    Dim mintJOrigen As Integer
    Dim mintROrigen As Integer
    Dim mintVOrigen As Integer

    Dim mintCodKilates As Integer

    Dim rsLocal As ADODB.Recordset

    Dim mblnFueraChange As Boolean

    Const cDigitos As String = "1234567890.,"
    Const cLetras As String = "BURLOSINEA.,"
    Dim intCodAlmacenOrigen As Integer
    Public mstrRuta As String
    Friend WithEvents btnBuscar As Button
    Public mstrArchivo As String

    'Public archivoImagenJMR As HttpPostedFileBase 

    '''Utiliza la variable global -gStrCodificacionImportes- para la codificacion de los costos en el abc de articulos
    Public Function Cifrar(ByRef cCantidad As String) As String
        Dim I As Integer 'Posición del dígito en la cadena cDigitos
        Dim c As Integer 'Contador de posiciones del parámetro
        Dim cDig As String 'Dígito o caracter extraído del parámetro cCantidad
        Dim cRes As String 'Cadena que va almacenando el resultado
        Dim cCant As String 'Almacena el parámetro sin espacios en blanco
        c = 0
        cRes = ""
        cCant = Trim(cCantidad)
        Do While c < Len(cCant)
            c = c + 1
            cDig = Mid(cCant, c, 1)
            I = InStr(1, cDigitos, cDig)
            If I <= 0 Then
                Cifrar = "( ERROR )"
                Exit Function
            End If
            cRes = cRes & Mid(gstrCodificacionImportes, I, 1)
        Loop
        Cifrar = cRes
        Return Cifrar
    End Function

    '''Utiliza la variable global -gStrCodificacionImportes- para la codificacion de los costos en el abc de articulos
    Public Function DesCifrar(ByRef cCantidadCifrada As String) As String
        Dim I As Integer 'Posición del caracter en la cadena cLetras
        Dim c As Integer 'Contador de posiciones del parámetro
        Dim cDig As String 'Dígito o caracter extraído del parámetro cCantidadCifrada
        Dim cRes As String 'Cadena que va almacenando el resultado
        Dim cCant As String 'Almacena el parámetro sin espacios en blanco
        c = 0
        cRes = ""
        cCant = Trim(cCantidadCifrada)
        Do While c < Len(cCant)
            c = c + 1
            cDig = Mid(cCant, c, 1)
            I = InStr(1, gstrCodificacionImportes, cDig)
            If I <= 0 Then
                DesCifrar = "( ERROR )"
                Exit Function
            End If
            cRes = cRes & Mid(cDigitos, I, 1)
        Loop
        DesCifrar = cRes
        Return DesCifrar
    End Function

    Public Sub ActualizaCantidades()
        txtCostoFactura(sstArticulo.SelectedIndex).Text = Cifrar(CStr(nCostoFactura))
        txtCostoAdicional(sstArticulo.SelectedIndex).Text = Cifrar(CStr(nCostoAdicional))
        txtCostoIndirecto(sstArticulo.SelectedIndex).Text = Cifrar(CStr(nCostoIndirectos))
        nCostoReal = nCostoFactura + nCostoAdicional + nCostoIndirectos
        txtCostoReal(sstArticulo.SelectedIndex).Text = Cifrar(CStr(nCostoReal))
    End Sub

    Sub Buscar()
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendrá el string del tag que se le mandara al fromulario de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas

        Dim cWHERE As String
        cWHERE = ""

        'strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mandó llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control


        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                cWHERE = IIf((mintJProv = 0), "", " And CodProveedor=  " & mintJProv & " ") & " And CodGrupo = " & gCODJOYERIA & " "
            Case nRELOJERIA
                cWHERE = IIf((mintRProv = 0), "", " And CodProveedor=  " & mintRProv & " ") & " And CodGrupo = " & gCODRELOJERIA & " "
            Case nVARIOS
                cWHERE = IIf((mintVProv = 0), "", " And CodProveedor=  " & mintVProv & " ") & " And CodGrupo = " & gCODVARIOS & " "
        End Select


        If UCase(strControlActual) = "TXTCODARTICULO" Then
            Select Case Me.sstArticulo.SelectedIndex
                Case nJOYERIA
                    strCaptionForm = "Consulta de Joyería"
                    gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, b.DescTipoMaterial AS MATERIAL, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a (Nolock) inner join CatTipoMaterial b (Nolock) on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODJOYERIA & " and a.CodArticulo >= " & CInt(Numerico((Me.txtCodArticulo.Text))) & " ORDER BY a.CodArticulo"
                Case nRELOJERIA
                    strCaptionForm = "Consulta de Relojería"
                    gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a (Nolock) inner join CatTipoMaterial b (Nolock) on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODRELOJERIA & " and a.CodArticulo >= " & CInt(Numerico((Me.txtCodArticulo.Text))) & " ORDER BY a.CodArticulo"
                Case nVARIOS
                    strCaptionForm = "Consulta de Artículos Varios"
                    gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, b.DescTipoMaterial AS MATERIAL, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a (Nolock) inner join CatTipoMaterial b (Nolock) on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODVARIOS & " and a.CodArticulo >= " & CInt(Numerico((Me.txtCodArticulo.Text))) & " ORDER BY a.CodArticulo"
                Case Else
                    'Sale de este sub para que no ejecute ninguna opción
                    Exit Sub
            End Select
        ElseIf UCase(strControlActual) = "TXTDESCARTICULO" Then
            Select Case Me.sstArticulo.SelectedIndex
                Case nJOYERIA
                    strCaptionForm = "Consulta de Joyería"
                    gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, b.DescTipoMaterial AS MATERIAL, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a (Nolock) inner join CatTipoMaterial b (Nolock) on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODJOYERIA & " and a.DescArticulo LIKE '" & Trim(Me.txtDescArticulo.Text) & "%' ORDER BY a.DescArticulo"
                Case nRELOJERIA
                    strCaptionForm = "Consulta de Relojería"
                    gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a (Nolock) inner join CatTipoMaterial b (Nolock) on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODRELOJERIA & " and a.DescArticulo LIKE '" & Trim(Me.txtDescArticulo.Text) & "%' ORDER BY a.DescArticulo"
                Case nVARIOS
                    strCaptionForm = "Consulta de Artículos Varios"
                    gStrSql = "select a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, b.DescTipoMaterial AS MATERIAL, a.CodigoArticuloProv AS 'COD. PROVEEDOR' from CatArticulos a (Nolock) inner join CatTipoMaterial b (Nolock) on a.CodTipoMaterial = b.CodTipoMaterial WHERE a.codGrupo = " & gCODVARIOS & " and a.DescArticulo LIKE '" & Trim(Me.txtDescArticulo.Text) & "%' ORDER BY a.DescArticulo"
                Case Else
                    'Sale de este sub para que no ejecute ninguna opción
                    Exit Sub
            End Select
        ElseIf UCase(strControlActual) = "TXTCODIGODELPROVEEDOR" Then
            Select Case Me.sstArticulo.SelectedIndex
                Case nJOYERIA
                    strCaptionForm = "Consulta de Joyería"
                    If cboAlmacen.SelectedValue(0).Text = "[ Vacío ... ]" Then
                        gStrSql = "SELECT     a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), OrigenAnt) " & "+ '-' + RIGHT(lTRIM(RTRIM(REPLICATE('0', 5) + CONVERT(CHAR(5), CodigoAnt))), 5) END AS [ANTERIOR], LTRIM(RTRIM(p.dESCPROVACREED)) " & "AS PROVEEDOR, a.CodigoArticuloProv AS 'ARTICULO PROV.',O.DescAlmacenOrigen AS 'ORIGEN' " & "FROM CatArticulos a (Nolock) INNER JOIN " & "CatTipoMaterial b (Nolock) ON a.CodTipoMaterial = b.CodTipoMaterial INNER JOIN " & "CATPROVACREED p (Nolock) ON A.CODPROVEEDOR = p.cODPROVACREED INNER JOIN CatOrigen O (Nolock) ON A.CodAlmacenOrigen = O.CodAlmacenOrigen " & "WHERE a.codGrupo = " & gCODJOYERIA & "   AND a.CodigoArticuloProv LIKE '" & Trim(_txtCodigodelProveedor_0.Text) & "%' " & cWHERE & "ORDER BY a.CodArticulo"
                    Else
                        gStrSql = "SELECT     a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), OrigenAnt) " & "+ '-' + RIGHT(lTRIM(RTRIM(REPLICATE('0', 5) + CONVERT(CHAR(5), CodigoAnt))), 5) END AS [ANTERIOR], LTRIM(RTRIM(p.dESCPROVACREED)) " & "AS PROVEEDOR, a.CodigoArticuloProv AS 'ARTICULO PROV.',O.DescAlmacenOrigen AS 'ORIGEN' " & "FROM CatArticulos a (Nolock) INNER JOIN " & "CatTipoMaterial b (Nolock) ON a.CodTipoMaterial = b.CodTipoMaterial INNER JOIN " & "CATPROVACREED p (Nolock) ON A.CODPROVEEDOR = p.cODPROVACREED INNER JOIN CatOrigen O (Nolock) ON A.CodAlmacenOrigen = O.CodAlmacenOrigen " & "WHERE a.codGrupo = " & gCODJOYERIA & "   AND a.CodigoArticuloProv LIKE '" & Trim(_txtCodigodelProveedor_0.Text) & "%' " & cWHERE & "AND A.CodAlmacenOrigen = " & mintJOrigen & " ORDER BY a.CodArticulo"
                    End If
                Case nRELOJERIA
                    strCaptionForm = "Consulta de Relojería"
                    If cboAlmacen.SelectedValue(1).Text = "[ Vacío ... ]" Then
                        gStrSql = "SELECT     a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), OrigenAnt) " & "+ '-' + RIGHT(lTRIM(RTRIM(REPLICATE('0', 5) + CONVERT(CHAR(5), CodigoAnt))), 5) END AS [ANTERIOR], LTRIM(RTRIM(p.dESCPROVACREED)) " & "AS PROVEEDOR, a.CodigoArticuloProv AS 'ARTICULO PROV.',O.DescAlmacenOrigen AS 'ORIGEN' " & "FROM CatArticulos a (Nolock) INNER JOIN " & "CatTipoMaterial b (Nolock) ON a.CodTipoMaterial = b.CodTipoMaterial INNER JOIN " & "CATPROVACREED p (Nolock) ON A.CODPROVEEDOR = p.cODPROVACREED INNER JOIN CatOrigen O (Nolock) ON A.CodAlmacenOrigen = O.CodAlmacenOrigen " & "WHERE a.codGrupo = " & gCODRELOJERIA & " AND a.CodigoArticuloProv LIKE '" & Trim(_txtCodigodelProveedor_1.Text) & "%' " & cWHERE & "ORDER BY a.CodArticulo"
                    Else
                        gStrSql = "SELECT     a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), OrigenAnt) " & "+ '-' + RIGHT(lTRIM(RTRIM(REPLICATE('0', 5) + CONVERT(CHAR(5), CodigoAnt))), 5) END AS [ANTERIOR], LTRIM(RTRIM(p.dESCPROVACREED)) " & "AS PROVEEDOR, a.CodigoArticuloProv AS 'ARTICULO PROV.',O.DescAlmacenOrigen AS 'ORIGEN' " & "FROM CatArticulos a (Nolock) INNER JOIN " & "CatTipoMaterial b (Nolock) ON a.CodTipoMaterial = b.CodTipoMaterial INNER JOIN " & "CATPROVACREED p (Nolock) ON A.CODPROVEEDOR = p.cODPROVACREED INNER JOIN CatOrigen O (Nolock) ON A.CodAlmacenOrigen = O.CodAlmacenOrigen " & "WHERE a.codGrupo = " & gCODRELOJERIA & "   AND a.CodigoArticuloProv LIKE '" & Trim(_txtCodigodelProveedor_1.Text) & "%' " & cWHERE & "AND A.CodAlmacenOrigen = " & mintROrigen & " ORDER BY a.CodArticulo"
                    End If
                Case nVARIOS
                    strCaptionForm = "Consulta de Artículos Varios"

                    If cboAlmacen.SelectedValue(2).Text = "[ Vacío ... ]" Then
                        gStrSql = "SELECT     a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), OrigenAnt) " & "+ '-' + RIGHT(lTRIM(RTRIM(REPLICATE('0', 5) + CONVERT(CHAR(5), CodigoAnt))), 5) END AS [ANTERIOR], LTRIM(RTRIM(p.dESCPROVACREED)) " & "AS PROVEEDOR, a.CodigoArticuloProv AS 'ARTICULO PROV.',O.DescAlmacenOrigen AS 'ORIGEN' " & "FROM CatArticulos a (Nolock) INNER JOIN " & "CatTipoMaterial b (Nolock) ON a.CodTipoMaterial = b.CodTipoMaterial INNER JOIN " & "CATPROVACREED p (Nolock) ON A.CODPROVEEDOR = p.cODPROVACREED INNER JOIN CatOrigen O (Nolock) ON A.CodAlmacenOrigen = O.CodAlmacenOrigen " & "WHERE a.codGrupo = " & gCODVARIOS & "   AND a.CodigoArticuloProv LIKE '" & Trim(_txtCodigodelProveedor_2.Text) & "%' " & cWHERE & "ORDER BY a.CodArticulo"
                    Else
                        gStrSql = "SELECT     a.CodArticulo AS CODIGO, LTrim(RTrim(a.DescArticulo)) AS DESCRIPCION, CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1), OrigenAnt) " & "+ '-' + RIGHT(lTRIM(RTRIM(REPLICATE('0', 5) + CONVERT(CHAR(5), CodigoAnt))), 5) END AS [ANTERIOR], LTRIM(RTRIM(p.dESCPROVACREED)) " & "AS PROVEEDOR, a.CodigoArticuloProv AS 'ARTICULO PROV.',O.DescAlmacenOrigen AS 'ORIGEN' " & "FROM CatArticulos a (Nolock) INNER JOIN " & "CatTipoMaterial b (Nolock) ON a.CodTipoMaterial = b.CodTipoMaterial INNER JOIN " & "CATPROVACREED p (Nolock) ON A.CODPROVEEDOR = p.cODPROVACREED INNER JOIN CatOrigen O (Nolock) ON A.CodAlmacenOrigen = O.CodAlmacenOrigen " & "WHERE a.codGrupo = " & gCODVARIOS & "   AND a.CodigoArticuloProv LIKE '" & Trim(_txtCodigodelProveedor_2.Text) & "%' " & cWHERE & "AND A.CodAlmacenOrigen = " & mintVOrigen & " ORDER BY a.CodArticulo"
                    End If
                Case Else
                    'Sale de este sub para que no ejecute ninguna opción
                    Exit Sub
            End Select
        End If

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
            MsgBox(C_msgSINDATOS & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If

        'Carga el formulario de consulta
        Dim FrmConsultas As FrmConsultas = New FrmConsultas()
        ConfiguraConsultas(FrmConsultas, 13150, RsGral, strTag, strCaptionForm)

        'If Me.sstArticulo.SelectedIndex = nRELOJERIA And strControlActual <> "TXTCODIGODELPROVEEDOR" Then
        '    ConfiguraConsultas(FrmConsultas, 8790, RsGral, strTag, strCaptionForm)
        'ElseIf strControlActual <> "TXTCODIGODELPROVEEDOR" Then
        '    ConfiguraConsultas(FrmConsultas, 10850, RsGral, strTag, strCaptionForm)
        'ElseIf strControlActual = "TXTCODIGODELPROVEEDOR" Then
        '    ConfiguraConsultas(FrmConsultas, 13150, RsGral, strTag, strCaptionForm)
        'End If

        With FrmConsultas.Flexdet
            Select Case strControlActual
                Case "TXTCODARTICULO", "TXTDESCARTICULO"
                    If Me.sstArticulo.SelectedIndex <> nRELOJERIA Then
                        .set_ColWidth(0, 0, 900) 'Columna del Código
                        .set_ColWidth(1, 0, 6000) 'Columna de la Descripción
                        .set_ColWidth(2, 0, 2055) 'Columna de Tipo de Material
                        .set_ColWidth(3, 0, 1890) 'Columna del Código del Proveedor
                        .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                        .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                        .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                        .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    Else
                        .set_ColWidth(0, 0, 900) 'Columna del Código
                        .set_ColWidth(1, 0, 6000) 'Columna de la Descripción
                        .set_ColWidth(2, 0, 1890) 'Columna del Código del Proveedor
                        .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                        .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                        .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    End If
                Case "TXTCODIGODELPROVEEDOR"
                    .set_ColWidth(0, 0, 900) 'Columna del Código
                    .set_ColWidth(1, 0, 4500) 'Columna de la Descripción
                    .set_ColWidth(2, 0, 1100) 'Columna del Codigo Anterior
                    .set_ColWidth(3, 0, 2355) 'Columna de la Descripcion del Proveedor
                    .set_ColWidth(4, 0, 1600) 'Columna del Codigo del Articulo de Proveedor
                    .set_ColWidth(5, 0, 2700) 'Columna de la Descripcion del Origen
                    .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                    .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
                    .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(4, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
                    .set_ColAlignment(5, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            End Select
        End With
        FrmConsultas.ShowDialog()
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub


    Sub BusquedaEspecial(ByRef CodArticulo As String)
        On Error GoTo Merr
        Dim strSQL As String
        Dim strTag As String 'Cadena que contendra el estring del tag que se le mandara al, fromularo de consultas
        Dim strCaptionForm As String 'Titulo que mostrara el formulario de consultas
        Dim strControlActual As String 'Nombre del control actual
        Dim Columna As Integer
        Dim cWHERE As String
        cWHERE = ""
        strControlActual = UCase(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name) 'Nombre del contro actual (Del que se mando llamar la consulta)
        strTag = UCase(Me.Name & "." & strControlActual) 'El tag sera el nombre del formulario + el nombre del control

        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                cWHERE = IIf((mintJProv = 0), "", " And CodProveedor=  " & mintJProv) & " And CodGrupo = " & gCODJOYERIA & " "
            Case nRELOJERIA
                cWHERE = IIf((mintRProv = 0), "", " And CodProveedor=  " & mintRProv) & " And CodGrupo = " & gCODRELOJERIA & " "
            Case nVARIOS
                cWHERE = IIf((mintVProv = 0), "", " And CodProveedor=  " & mintVProv) & " And CodGrupo = " & gCODVARIOS & " "
        End Select

        strCaptionForm = "Consulta de Articulos"
        '        If UCase(Me.ActiveControl.Name) <> "TXTCODARTICULO" Then
        '        If UCase(Me.ActiveControl.Name) <> "TXTDESCARTICULO"
        If strControlActual = "TXTCODARTICULO" Or strControlActual = "TXTDESCARTICULO" Then
            strSQL = "SELECT     CodArticulo AS CODIGO, RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION, " & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR], " & "dbo.FormatCantidad(A.PrecioPubDolar)  AS [PRECIO PÚBLICO] , " & "case PesosFijos WHEN 0 THEN 'DÓLARES' WHEN 1 THEN 'PESOS' END AS [MONEDA] " & "From CatArticulos A cross Join Configuraciongeneral c WHERE (CodArticulo = " & Trim(CodArticulo) & ") " & "OR   (OrigenAnt = " & CInt((CodArticulo)) & ") AND (CodigoAnt = " & CInt((CodArticulo)) & ")"
        Else
            strTag = "FRMCORPOABCARTICULOS.TXTCODIGODELPROVEEDOR"
            strSQL = "SELECT     CodArticulo AS CODIGO, RTRIM(LTRIM(DescArticulo)) AS DESCRIPCION, " & "CASE CodigoAnt WHEN 0 THEN '' ELSE CONVERT(CHAR(1),OrigenAnt)+'-'+RIGHT( lTRIM(RTRIM( REPLICATE('0',5)+ CONVERT(CHAR(5),CodigoAnt))) ,5) End as [CODIGO ANTERIOR], " & "dbo.FormatCantidad(A.PrecioPubDolar)  AS [PRECIO PÚBLICO] , " & "case PesosFijos WHEN 0 THEN 'DÓLARES' WHEN 1 THEN 'PESOS' END AS [MONEDA] " & "From CatArticulos A cross Join Configuraciongeneral c WHERE CodigoArticuloProv = '" & CodArticulo & "' " & cWHERE

        End If
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, strSQL))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta entonces manda mensage y sale del procedimiento
        If RsGral.RecordCount = 0 Then
            MsgBox("El Artículo no existe." & vbNewLine & "Verifique por favor", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            RsGral.Close()
            Exit Sub
        End If

        'Carga el formulario de consulta

        'Load(FrmConsultas)
        'ModVariables.frmConsultas.Show()
        ConfiguraConsultas(FrmConsultas, 11050, RsGral, strTag, strCaptionForm)

        With FrmConsultas.Flexdet
            .set_ColWidth(0, 0, 900)
            .set_ColWidth(1, 0, 4800)
            .set_ColWidth(2, 0, 1900)
            .set_ColWidth(3, 0, 1620)
            .set_ColWidth(4, 0, 1800)

            .set_ColAlignment(0, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColAlignment(1, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            .set_ColAlignment(2, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignLeftCenter)
            .set_ColAlignment(3, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignRightCenter)
            .set_ColAlignment(4, MSHierarchicalFlexGridLib.AlignmentSettings.flexAlignCenterCenter)

            .WordWrap = False
        End With
        mblnFueraChange = True
        CentrarForma(FrmConsultas)
        FrmConsultas.ShowDialog()
        mblnFueraChange = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Sub Eliminar()
        On Error GoTo Merr
        Dim blnTransaction As Boolean

        If BuscaArticulo(mintCodArticulo) Then
            'Pregunta por la integridad referencial
            If ModCorporativo.Referencia("SELECT codArticulo FROM OrdenesCompraPreCat WHERE CodArticulo = " & mintCodArticulo) Then
                MsgBox("No puede borrar el registro debido a su existencia en una ó más Órdenes de Compra", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Me.txtCodArticulo.Focus()
                ModEstandar.SelTxt()
                Exit Sub
            ElseIf ModCorporativo.Referencia("SELECT codArticulo FROM Inventario WHERE CodArticulo = " & mintCodArticulo) Then
                MsgBox("No puede borrar el registro debido a su existencia en el inventario", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Me.txtCodArticulo.Focus()
                ModEstandar.SelTxt()
                Exit Sub
            End If

            If MsgBox(C_msgBORRAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa) <> MsgBoxResult.Yes Then
                Me.txtCodArticulo.Focus()
                ModEstandar.SelTextoTxt((Me.txtCodArticulo))
                Exit Sub
            End If
            Cnn.BeginTrans()
            blnTransaction = True
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            Select Case Me.sstArticulo.SelectedIndex
                Case nJOYERIA
                    ModStoredProcedures.PR_IMECatArticulos(CStr(mintCodArticulo), Trim(Me.txtDescripcion(nJOYERIA).Text), CStr(gCODJOYERIA), CStr(mintJFam), CStr(mintJLin), CStr(mintJSub), CStr(mintCodKilates), CStr(0), CStr(0), CStr(mintJMaterial), Trim(""), Trim(""), CStr(lCrono), CStr(mintJUnidad), CStr(mintJOrigen), CStr(mintJProv), Trim(Me.txtCodigodelProveedor(nJOYERIA).Text), Trim(cMonedaCompra), Trim(Me.txtPrecioenDolares(nJOYERIA).Text), CStr(nCostoFactura), CStr(nCostoAdicional), CStr(nCostoIndirectos), CStr(nCostoReal), CStr(nCostoFacturaPesos), CStr(nCostoAdicionalPesos), CStr(nCostoIndirectosPesos), CStr(0), CStr(0), CStr(0), Trim(txtAdicional(0).Text), CStr(0), "", "", "", C_ELIMINACION, CStr(0)) '''27OCT2010 - MAVF
                Case nRELOJERIA
                    ModStoredProcedures.PR_IMECatArticulos(CStr(mintCodArticulo), Trim(Me.txtDescripcion(nRELOJERIA).Text), CStr(gCODRELOJERIA), CStr(0), CStr(0), CStr(0), CStr(mintCodKilates), CStr(mintRMar), CStr(mintRMod), CStr(mintRMaterial), Trim(cGenero), Trim(cMovimiento), CStr(lCrono), CStr(mintRUnidad), CStr(mintROrigen), CStr(mintRProv), Trim(Me.txtCodigodelProveedor(nRELOJERIA).Text), Trim(cMonedaCompra), Trim(Me.txtPrecioenDolares(nRELOJERIA).Text), CStr(nCostoFactura), CStr(nCostoAdicional), CStr(nCostoIndirectos), CStr(nCostoReal), CStr(nCostoFacturaPesos), CStr(nCostoAdicionalPesos), CStr(nCostoIndirectosPesos), CStr(0), CStr(0), CStr(0), Trim(txtAdicional(1).Text), CStr(0), "", "", "", C_ELIMINACION, CStr(0)) '''27OCT2010 - MAVF
                Case nVARIOS
                    ModStoredProcedures.PR_IMECatArticulos(CStr(mintCodArticulo), Trim(Me.txtDescripcion(nVARIOS).Text), CStr(gCODVARIOS), CStr(mintVFam), CStr(mintVLin), CStr(0), CStr(0), CStr(0), CStr(0), CStr(mintVMaterial), Trim(""), Trim(""), CStr(False), CStr(mintVUnidad), Str(mintVOrigen), CStr(mintVProv), Trim(Me.txtCodigodelProveedor(nVARIOS).Text), Trim(cMonedaCompra), Trim(Me.txtPrecioenDolares(nVARIOS).Text), CStr(nCostoFactura), CStr(nCostoAdicional), CStr(nCostoIndirectos), CStr(nCostoReal), CStr(nCostoFacturaPesos), CStr(nCostoAdicionalPesos), CStr(nCostoIndirectosPesos), CStr(0), CStr(0), CStr(0), Trim(txtAdicional(2).Text), CStr(0), "", "", "", C_ELIMINACION, CStr(0)) '''27OCT2010 - MAVF
            End Select

            Cmd.Execute()
            Cnn.CommitTrans()
            blnTransaction = False
        Else
            MsgBox("Proporcione un código válido para Borrar", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            RsGral.Close()
            Exit Sub
        End If

        Me.txtDescArticulo.Text = ""
        Me.txtDescArticulo.Tag = ""
        mintCodArticulo = 0
        mblnNuevo = True
        Limpiar()
        Me.txtDescArticulo.Focus()

Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Function Guardar() As Boolean
        On Error GoTo Merr
        Dim blnTransaction As Boolean
        Dim PesosFijos As Byte
        Dim OrigenAnt As Integer
        Dim CodigoAnt As Integer
        Dim Archivo As String
        Dim Extension As String
        Dim I As Object
        Dim Contador As Integer
        Dim miArchivo As String

        'If Not Cambios() Then
        '    Limpiar()
        '    Exit Function
        'End If

        '''27OCT2010 - MAVF
        If Not ValidaDatosManejoDiamanteSuelto() Then
            If CInt(Numerico((txtCodArticulo.Text))) = 0 Then
                mblnNuevo = True
            End If
            Exit Function
        End If

        'Valida si todos los datos son válidos
        If Not ValidaDatos() Then
            If CInt(Numerico((txtCodArticulo.Text))) = 0 Then
                mblnNuevo = True
            End If
            Exit Function
        End If

        Cnn.BeginTrans()
        blnTransaction = True
        'Calcular  los costos en pesos
        nCostoAdicionalPesos = nCostoAdicional * gcurCorpoTIPOCAMBIODOLAR
        nCostoFacturaPesos = nCostoFactura * gcurCorpoTIPOCAMBIODOLAR
        nCostoIndirectosPesos = nCostoIndirectos * gcurCorpoTIPOCAMBIODOLAR

        If chkCodigoAnterior.CheckState = System.Windows.Forms.CheckState.Checked Then
            OrigenAnt = intCodAlmacenOrigen
            CodigoAnt = CInt(Numerico(txtCodArtAnterior.Text))
        Else
            OrigenAnt = 0
            CodigoAnt = 0
        End If

        If mblnNuevo Then
            'Añadir
            Select Case Me.sstArticulo.SelectedIndex
                Case nJOYERIA
                    PesosFijos = IIf((_optMoneda_10.Checked = True), 0, 1)
                    ModStoredProcedures.PR_IMECatArticulos(CStr(mintCodArticulo), Trim(Me._txtDescripcion_0.Text), CStr(gCODJOYERIA), CStr(mintJFam), CStr(mintJLin), CStr(mintJSub), CStr(mintCodKilates), CStr(0), CStr(0), CStr(mintJMaterial), Trim(""), Trim(""), CStr(lCrono), CStr(mintJUnidad), CStr(mintJOrigen), CStr(mintJProv), Trim(Me._txtCodigodelProveedor_0.Text), Trim(cMonedaCompra), Trim(Me._txtPrecioenDolares_0.Text), CStr(nCostoFactura), CStr(nCostoAdicional), CStr(nCostoIndirectos), CStr(nCostoReal), CStr(nCostoFacturaPesos), CStr(nCostoAdicionalPesos), CStr(nCostoIndirectosPesos), CStr(PesosFijos), CStr(OrigenAnt), CStr(CodigoAnt), Trim(_txtCostoAdicional_0.Text), Trim(CStr(ModEstandar.Numerico((txtMDSPeso.Text)))), Trim(txtMDSColor.Text), Trim(txtMDSPureza.Text), Trim(txtMDSCertificado.Text), C_INSERCION, CStr(0)) '''27OCT2010 - MAVF
                Case nRELOJERIA
                    PesosFijos = IIf((optMoneda(7).Checked = True), 0, 1)
                    ModStoredProcedures.PR_IMECatArticulos(CStr(mintCodArticulo), Trim(Me._txtDescripcion_1.Text), CStr(gCODRELOJERIA), CStr(0), CStr(0), CStr(0), CStr(mintCodKilates), CStr(mintRMar), CStr(mintRMod), CStr(mintRMaterial), Trim(cGenero), Trim(cMovimiento), CStr(lCrono), CStr(mintRUnidad), CStr(mintROrigen), CStr(mintRProv), Trim(Me._txtCodigodelProveedor_1.Text), Trim(cMonedaCompra), Trim(Me._txtPrecioenDolares_1.Text), CStr(nCostoFactura), CStr(nCostoAdicional), CStr(nCostoIndirectos), CStr(nCostoReal), CStr(nCostoFacturaPesos), CStr(nCostoAdicionalPesos), CStr(nCostoIndirectosPesos), CStr(PesosFijos), CStr(OrigenAnt), CStr(CodigoAnt), Trim(_txtCostoAdicional_1.Text), CStr(0), "", "", "", C_INSERCION, CStr(0))
                Case nVARIOS
                    PesosFijos = IIf((_optMoneda_8.Checked = True), 0, 1)
                    ModStoredProcedures.PR_IMECatArticulos(CStr(mintCodArticulo), Trim(Me._txtDescripcion_2.Text), CStr(gCODVARIOS), CStr(mintVFam), CStr(mintVLin), CStr(0), CStr(0), CStr(0), CStr(0), CStr(mintVMaterial), Trim(""), Trim(""), CStr(False), CStr(mintVUnidad), Str(mintVOrigen), CStr(mintVProv), Trim(Me._txtCodigodelProveedor_2.Text), Trim(cMonedaCompra), Trim(Me._txtPrecioenDolares_2.Text), CStr(nCostoFactura), CStr(nCostoAdicional), CStr(nCostoIndirectos), CStr(nCostoReal), CStr(nCostoFacturaPesos), CStr(nCostoAdicionalPesos), CStr(nCostoIndirectosPesos), CStr(PesosFijos), CStr(OrigenAnt), CStr(CodigoAnt), Trim(_txtCostoAdicional_2.Text), CStr(0), "", "", "", C_INSERCION, CStr(0))
            End Select
            Cmd.Execute()
            mintCodArticulo = Cmd.Parameters("ID").Value
        Else
            'Modificar
            Select Case Me.sstArticulo.SelectedIndex
                Case nJOYERIA
                    If mintJFam = 0 Then
                        Me._txtDescripcion_0.Text = Me._lblDescripcion_0.Text
                    End If
                    PesosFijos = IIf((_optMoneda_10.Checked = True), 0, 1)
                    ModStoredProcedures.PR_IMECatArticulos(CStr(Me.txtCodArticulo.Text), Trim(Me._txtDescripcion_0.Text), CStr(gCODJOYERIA), CStr(mintJFam), CStr(mintJLin), CStr(mintJSub), CStr(mintCodKilates), CStr(0), CStr(0), CStr(mintJMaterial), Trim(""), Trim(""), CStr(lCrono), CStr(mintJUnidad), CStr(mintJOrigen), CStr(mintJProv), Trim(Me._txtCodigodelProveedor_0.Text), Trim(cMonedaCompra), Trim(Me._txtPrecioenDolares_0.Text), CStr(nCostoFactura), CStr(nCostoAdicional), CStr(nCostoIndirectos), CStr(nCostoReal), CStr(nCostoFacturaPesos), CStr(nCostoAdicionalPesos), CStr(nCostoIndirectosPesos), CStr(PesosFijos), CStr(OrigenAnt), CStr(CodigoAnt), Trim(_txtAdicional_0.Text), Trim(CStr(ModEstandar.Numerico((txtMDSPeso.Text)))), Trim(txtMDSColor.Text), Trim(txtMDSPureza.Text), Trim(txtMDSCertificado.Text), C_MODIFICACION, CStr(0)) '''27OCT2010 - MAVF
                Case nRELOJERIA
                    If mintRMar = 0 Then
                        Me._txtDescripcion_1.Text = Me._lblDescripcion_1.Text
                    End If
                    PesosFijos = IIf((optMoneda(7).Checked = True), 0, 1)
                    ModStoredProcedures.PR_IMECatArticulos(Trim(Me.txtCodArticulo.Text), Trim(Me._txtDescripcion_1.Text), CStr(gCODRELOJERIA), CStr(0), CStr(0), CStr(0), CStr(mintCodKilates), CStr(mintRMar), CStr(mintRMod), CStr(mintRMaterial), Trim(cGenero), Trim(cMovimiento), CStr(lCrono), CStr(mintRUnidad), CStr(mintROrigen), CStr(mintRProv), Trim(Me._txtCodigodelProveedor_1.Text), Trim(cMonedaCompra), Trim(Me._txtPrecioenDolares_1.Text), CStr(nCostoFactura), CStr(nCostoAdicional), CStr(nCostoIndirectos), CStr(nCostoReal), CStr(nCostoFacturaPesos), CStr(nCostoAdicionalPesos), CStr(nCostoIndirectosPesos), CStr(PesosFijos), CStr(OrigenAnt), CStr(CodigoAnt), Trim(_txtAdicional_1.Text), CStr(0), "", "", "", C_MODIFICACION, CStr(0))
                Case nVARIOS
                    If mintVFam = 0 Then
                        Me._txtDescripcion_2.Text = Me._lblDescripcion_2.Text
                    End If
                    PesosFijos = IIf((_optMoneda_8.Checked = True), 0, 1)
                    ModStoredProcedures.PR_IMECatArticulos(Trim(Me.txtCodArticulo.Text), Trim(Me._txtDescripcion_2.Text), CStr(gCODVARIOS), CStr(mintVFam), CStr(mintVLin), CStr(0), CStr(0), CStr(0), CStr(0), CStr(mintVMaterial), Trim(""), Trim(""), CStr(False), CStr(mintVUnidad), Str(mintVOrigen), CStr(mintVProv), Trim(Me._txtCodigodelProveedor_2.Text), Trim(cMonedaCompra), Trim(Me._txtPrecioenDolares_2.Text), CStr(nCostoFactura), CStr(nCostoAdicional), CStr(nCostoIndirectos), CStr(nCostoReal), CStr(nCostoFacturaPesos), CStr(nCostoAdicionalPesos), CStr(nCostoIndirectosPesos), CStr(PesosFijos), CStr(OrigenAnt), CStr(CodigoAnt), Trim(_txtAdicional_2.Text), CStr(0), "", "", "", C_MODIFICACION, CStr(0))
            End Select
            Cmd.Execute()
        End If

        txtCodArticulo.Text = CStr(mintCodArticulo)
        'Guardar el archivo de imagen en el servidor
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                Archivo = Trim(txtImagen(nJOYERIA).Text)
            Case nRELOJERIA
                Archivo = Trim(txtImagen(nRELOJERIA).Text)
            Case nVARIOS
                Archivo = Trim(txtImagen(nVARIOS).Text)
        End Select
        If Archivo <> "" Then
            Extension = ""
            Contador = 0
            For I = Len(Archivo) To 1 Step -1
                If Mid(Archivo, I, 1) <> "." Then
                    Contador = Contador + 1
                Else
                    Extension = Mid(Archivo, I + 1, Contador)
                    Exit For
                End If
            Next

            'Determinar si existe el archivo que estamos dando
            miArchivo = Dir(My.Application.Info.DirectoryPath & "\Sistema\Imagenes\" & Trim(txtCodArticulo.Text) & "." & Extension)
            If miArchivo <> "" Then
                Kill(My.Application.Info.DirectoryPath & "\Sistema\Imagenes\" & Trim(txtCodArticulo.Text) & "." & Extension)
            End If
            FileCopy(Archivo, My.Application.Info.DirectoryPath & "\Sistema\Imagenes\" & Trim(txtCodArticulo.Text) & "." & Extension)
        End If

        Cnn.CommitTrans()
        blnTransaction = False
        If mblnNuevo Then
            MsgBox("El Artículo ha sido grabado correctamente con el código " & mintCodArticulo, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
        End If
        Nuevo()
        mblnNuevo = True
        Limpiar()
        Guardar = True
        txtCodArticulo.Focus()

Merr:
        If Err.Number <> 0 Then
            If blnTransaction Then Cnn.RollbackTrans()
            ModEstandar.MostrarError()
        End If
    End Function

    Function ObtenerCodArticulodeCodProv(ByRef CodProveedor As String) As String
        ObtenerCodArticulodeCodProv = ""
        gStrSql = "Select * from CatArticulos Where CodigoArticuloProv = '" & CodProveedor & "'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute
        Select Case RsGral.RecordCount
            Case 0
                ObtenerCodArticulodeCodProv = CStr(-1)
            Case 1
                ObtenerCodArticulodeCodProv = RsGral.Fields("CodArticulo").Value
            Case Is > 1
                'mas de un articulo
                ObtenerCodArticulodeCodProv = CStr(-2)
        End Select
    End Function

    Private Function ValidaDatosManejoDiamanteSuelto() As Boolean
        Dim vlResult As Byte

        Select Case sstArticulo.SelectedIndex
            Case nJOYERIA
                lblEstatus.Text = ""
                If Trim(txtMDSPeso.Text) = "" Or CDec(ModEstandar.Numerico((txtMDSPeso.Text))) = 0 Then lblEstatus.Text = "CT-" '''08NOV2010 - MAVF
                If Trim(txtMDSColor.Text) = "" Then lblEstatus.Text = lblEstatus.Text & "COLOR-"
                If Trim(txtMDSPureza.Text) = "" Then lblEstatus.Text = lblEstatus.Text & "Q-"
                ''' 27OCT2010 - MAVF
                ''' Si no estan capturados todos los datos o si falta alguno, entonces notifica y pregunta si kiere capturarlossss
                If Trim(lblEstatus.Text) <> "" Then
                    lblEstatus.Text = Mid(lblEstatus.Text, 1, Len(lblEstatus.Text) - 1)
                    vlResult = MsgBox("Datos de Diamante Suelto no capturados..." & vbNewLine & "Desea ingresarlos ???", MsgBoxStyle.Information + MsgBoxStyle.YesNo, gstrNombCortoEmpresa)
                    If vlResult = MsgBoxResult.Yes Then
                        txtMDSPeso.Focus()
                        ValidaDatosManejoDiamanteSuelto = False '''Si kiere capturarlos
                    Else
                        ValidaDatosManejoDiamanteSuelto = True '''No kiere capturarlos
                    End If
                Else
                    ValidaDatosManejoDiamanteSuelto = True '''Todos los datos estan capturados
                End If

            Case Else
                ValidaDatosManejoDiamanteSuelto = True '''Es de Relojeria y de Varios
        End Select

    End Function

    Public Function ValidaDatos() As Boolean
        Dim lFam As String
        Dim lLin As String
        Dim lSubL As String
        Dim lKil As String
        Dim lTipoM As String
        Dim lMar As String
        Dim lMod As String
        Dim lGen As String
        Dim lMov As String
        Dim lCrono As String
        Dim lDescGpo As String

        ValidaDatos = True
        lFam = "" : lLin = "" : lSubL = "" : lKil = "" : lTipoM = ""
        lMar = "" : lMod = "" : lGen = "" : lMov = "" : lCrono = "" : lTipoM = ""
        lFam = "" : lLin = "" : lTipoM = ""
        lDescGpo = ""

        Select Case sstArticulo.SelectedIndex
            Case nJOYERIA
                If mintJFam = 0 Then
                    MsgBox("Indique la familia a la que pertenece el artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _dbcFamilia_0.Focus()
                    ValidaDatos = False
                ElseIf mintCodKilates = 0 Then
                    ''' 08NOV2010 - MAVF -SE REABRIO VALIDACION PARA CONSIDERAR SOLO SIN KILTAJE COMO CLASIFICACIÓN NULA DE KILATES
                    ''' 26MAR2008 - MAVF
                    MsgBox("Indique el kilataje del artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    dbcKilates.Focus()
                    ValidaDatos = False
                ElseIf mintJMaterial = 0 Then
                    MsgBox("Indique el Tipo de Material del que está hecho el Artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _dbcMaterial_0.Focus()
                    ValidaDatos = False
                ElseIf CDec(ModEstandar.Numerico(_txtPrecioenDolares_0.Text)) < 0 Then
                    MsgBox("Especifique el precio al Público en Dólares", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtPrecioenDolares_0.Focus()
                    ValidaDatos = False
                ElseIf nCostoFactura < 0 Then
                    MsgBox("Especifique el Costo Factura del Artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtCostoFactura_0.Focus()
                    ValidaDatos = False
                ElseIf nCostoAdicional < 0 Then
                    MsgBox("Especifique el Costo Adicional del Artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtCostoAdicional_0.Focus()
                    ValidaDatos = False
                ElseIf nCostoIndirectos < 0 Then
                    MsgBox("Especifique el Costo Indirecto del Artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtCostoIndirecto_0.Focus()
                    ValidaDatos = False
                ElseIf mintJUnidad = 0 Then
                    MsgBox("Indique la Unidad de Presentación del artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _cboUnidad_0.Focus()
                    ValidaDatos = False

                ElseIf mintJOrigen = 0 And (_cboAlmacen_0.Text = "" Or _cboAlmacen_0.Text = cINDEFINIDO) Then
                    MsgBox("Indique el Almacén de Origen al que pertenece el artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _cboAlmacen_0.Focus()
                    ValidaDatos = False
                ElseIf mintJProv = 0 Then
                    MsgBox("Indique la descripción del Proveedor del artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _dbcProveedor_0.Focus()
                    ValidaDatos = False
                Else
                    ValidaDatos = True
                End If

                lDescGpo = "Joyería"
                DefineCondicionesJoy(lFam, lLin, lSubL, lKil, lTipoM)
                gStrSql = "SELECT * FROM CatArticulos (Nolock) WHERE CodProveedor = " & mintJProv & " And " & lFam & " And " & lLin & " And " & lSubL & " And " & lKil & " And " & lTipoM & " And ltrim(rtrim(Adicional)) = '" & Trim(_txtAdicional_0.Text) & "' And ltrim(rtrim(CodigoArticuloProv)) = '" & Trim(_txtCodigodelProveedor_0.Text) & "' "

            Case nRELOJERIA
                If mintRMar = 0 Then
                    MsgBox("Indique la Marca del Reloj", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    dbcMarca.Focus()
                    ValidaDatos = False
                ElseIf mintRMaterial = 0 Then
                    MsgBox("Indique el Tipo de Material del que está formado el Reloj", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _dbcMaterial_1.Focus()
                    ValidaDatos = False
                ElseIf CDec(ModEstandar.Numerico(_txtPrecioenDolares_1.Text)) < 0 Then
                    MsgBox("Especifique el precio al Público en Dólares", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtPrecioenDolares_1.Focus()
                    ValidaDatos = False
                ElseIf nCostoFactura < 0 Then
                    MsgBox("Especifique el Costo Factura del Reloj", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtCostoFactura_1.Focus()
                    ValidaDatos = False
                ElseIf nCostoAdicional < 0 Then
                    MsgBox("Especifique el Costo Adicional del Reloj", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtCostoAdicional_1.Focus()
                    ValidaDatos = False
                ElseIf nCostoIndirectos < 0 Then
                    MsgBox("Especifique el Costo Indirecto del Reloj", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtCostoIndirecto_1.Focus()
                    ValidaDatos = False
                ElseIf nCostoReal < 0 Then
                    MsgBox("Especifique el Costo Real del Reloj", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtCostoReal_1.Focus()
                    ValidaDatos = False
                ElseIf mintRUnidad = 0 Then
                    MsgBox("Indique la Unidad de Presentación del Reloj o los Relojes", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _cboUnidad_1.Focus()
                    ValidaDatos = False

                ElseIf mintROrigen = 0 And (_cboAlmacen_1.Text = "" Or _cboAlmacen_1.Text = cINDEFINIDO) Then
                    MsgBox("Especifique el Almacén de Origen al que pertenece el Reloj", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _cboAlmacen_1.Focus()
                    ValidaDatos = False
                ElseIf mintRProv = 0 Then
                    MsgBox("Indique la descripción del Proveedor del artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _dbcProveedor_1.Focus()
                    ValidaDatos = False
                Else
                    ValidaDatos = True
                End If

                lDescGpo = "Relojería"
                DefineCondicionesRel(lMar, lMod, lGen, lMov, lCrono, lTipoM)
                gStrSql = " SELECT * FROM CatArticulos (Nolock) WHERE CodProveedor = " & mintRProv & " And " & lMar & " And " & lMod & " And " & lGen & " And " & lMov & " And " & lCrono & " And " & lTipoM & " And ltrim(rtrim(Adicional)) = '" & Trim(_txtAdicional_1.Text) & "' And ltrim(rtrim(CodigoArticuloProv)) = '" & Trim(_txtCodigodelProveedor_1.Text) & "' "

            Case nVARIOS
                If mintVFam = 0 Then
                    MsgBox("Especifique la familia a la que pertenece el artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _dbcFamilia_1.Focus()
                    ValidaDatos = False
                ElseIf mintVMaterial = 0 Then
                    MsgBox("Indique el Tipo de Material del que está hecho el Artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _dbcMaterial_2.Focus()
                    ValidaDatos = False
                ElseIf CDec(Numerico(_txtPrecioenDolares_2.Text)) < 0 Then
                    MsgBox("Especifique el precio al Público en Dólares", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtPrecioenDolares_2.Focus()
                    ValidaDatos = False
                ElseIf nCostoFactura < 0 Then
                    MsgBox("Especifique el Costo Factura del Artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtCostoFactura_2.Focus()
                    ValidaDatos = False
                ElseIf nCostoAdicional < 0 Then
                    MsgBox("Especifique el Costo Adicional del Artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtCostoAdicional_2.Focus()
                    ValidaDatos = False
                ElseIf nCostoIndirectos < 0 Then
                    MsgBox("Especifique el Costo Indirecto del Artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _txtCostoIndirecto_2.Focus()
                    ValidaDatos = False
                ElseIf mintVUnidad = 0 Then
                    MsgBox("Indique la Unidad de Presentación del Artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _cboUnidad_2.Focus()
                    ValidaDatos = False

                ElseIf mintVOrigen = 0 And (_cboAlmacen_2.Text = "" Or _cboAlmacen_2.Text = cINDEFINIDO) Then
                    MsgBox("Especifique el Almacén de Origen al que pertenece el artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _cboAlmacen_2.Focus()
                    ValidaDatos = False
                ElseIf mintVProv = 0 Then
                    MsgBox("Indique la descripción del Proveedor del artículo", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
                    _dbcProveedor_2.Focus()
                    ValidaDatos = False
                Else
                    ValidaDatos = True
                End If

                lDescGpo = "Varios"
                DefineCondicionesVar(lFam, lLin, lTipoM)
                gStrSql = "SELECT * FROM CatArticulos (Nolock) WHERE CodProveedor = " & mintVProv & " And " & lFam & " And " & lLin & " And " & lTipoM & " And ltrim(rtrim(Adicional)) = '" & Trim(_txtAdicional_2.Text) & "' And ltrim(rtrim(CodigoArticuloProv)) = '" & Trim(_txtCodigodelProveedor_2.Text) & "' "

        End Select

        If chkCodigoAnterior.CheckState Then

            If Trim(dbcOrigen.Text) <> "" Then
                If Trim(txtCodArtAnterior.Text) = "" Then
                    MsgBox("Debe indicar el codigo anterior del artículo" & vbNewLine & "Favor de verificar..", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                    ValidaDatos = False
                    txtCodArtAnterior.Focus()
                    Exit Function
                End If
            Else
                MsgBox("Debe indicar el origen anterior del artículo" & vbNewLine & "Favor de verificar..", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
                ValidaDatos = False
                dbcOrigen.Focus()
                Exit Function
            End If
        End If

        If ValidaDatos Then

            If mblnNuevo Then '''si es nuevo  entonces ya existe otro con esa clasificación
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.UP_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                rsLocal = Cmd.Execute

                If rsLocal.RecordCount > 0 Then

                    If MsgBox("Ya existe un artículo con esta clasificación" & vbNewLine & "para el grupo de " & lDescGpo & vbNewLine & "para el Prov:  " & Trim(dbcProveedor.SelectedValue(sstArticulo.SelectedIndex).Text) & vbNewLine & vbNewLine & "Desea registrarlo de cualquier manera???", MsgBoxStyle.Information + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA) = MsgBoxResult.Yes Then
                        ValidaDatos = True
                    Else
                        ValidaDatos = False
                    End If
                Else
                    ValidaDatos = True
                End If
            Else '''si no es nuevo busca otros articulos con la clasificacion indicada actualmente

                gStrSql = gStrSql & " And CodArticulo <> " & CInt(Numerico((txtCodArticulo.Text)))
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.UP_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
                rsLocal = Cmd.Execute

                If Not rsLocal.EOF Then
                    If rsLocal.RecordCount > 0 Then

                        If MsgBox("Ya existe uno o varios artículo con esta clasificación" & vbNewLine & "para el grupo de " & lDescGpo & vbNewLine & "para el Prov:  " & Trim(dbcProveedor.SelectedValue(sstArticulo.SelectedIndex).Text) & vbNewLine & vbNewLine & "Desea registrarlo de cualquier manera???", MsgBoxStyle.Information + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA) = MsgBoxResult.Yes Then
                            ValidaDatos = True
                        Else
                            ValidaDatos = False
                        End If
                    End If
                Else
                    ValidaDatos = True
                End If
            End If
        End If

    End Function

    Public Function Cambios() As Boolean
        'Select Case Me.sstArticulo.SelectedIndex
        '    Case nJOYERIA

        '        If Trim(Me.dbcFamilia.SelectedValue(0).Text) <> Trim(Me.dbcFamilia.SelectedValue(0).Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.dbcLinea.SelectedValue(0).Text) <> Trim(Me.dbcLinea.SelectedValue(0).Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.dbcSubLinea.Text) <> Trim(Me.dbcSubLinea.Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.dbcKilates.Text) <> Trim(Me.dbcKilates.Tag) Then
        '            Cambios = True
        '        ElseIf Trim(Me.txtDescripcion(nJOYERIA).Text) <> Trim(Me.txtDescripcion(nJOYERIA).Tag) Then
        '            Cambios = True
        '        ElseIf Trim(Me.txtMDSPeso.Text) <> Trim(Me.txtMDSPeso.Tag) Then  '''27OCT2010 - MAVF
        '            Cambios = True
        '        ElseIf Trim(Me.txtMDSColor.Text) <> Trim(Me.txtMDSColor.Tag) Then  '''27OCT2010 - MAVF
        '            Cambios = True
        '        ElseIf Trim(Me.txtMDSPureza.Text) <> Trim(Me.txtMDSPureza.Tag) Then  '''27OCT2010 - MAVF
        '            Cambios = True
        '        ElseIf Trim(Me.txtMDSCertificado.Text) <> Trim(Me.txtMDSCertificado.Tag) Then  '''27OCT2010 - MAVF
        '            Cambios = True
        '        ElseIf Trim(cMonedaCompra) <> Trim(cMonedaCompraTag) Then
        '            Cambios = True
        '        ElseIf ModEstandar.Numerico(Me.txtPrecioenDolares(nJOYERIA).Text) <> ModEstandar.Numerico(Me.txtPrecioenDolares(nJOYERIA).Tag) Then
        '            Cambios = True
        '        ElseIf nCostoFactura <> nCostoFacturaTag Then
        '            Cambios = True
        '        ElseIf nCostoAdicional <> nCostoAdicionalTag Then
        '            Cambios = True
        '        ElseIf nCostoIndirectos <> nCostoIndirectosTag Then
        '            Cambios = True
        '        ElseIf nCostoReal <> nCostoRealTag Then
        '            Cambios = True

        '        ElseIf Trim(Me.dbcMaterial.SelectedValue(nJOYERIA).Text) <> Trim(Me.dbcMaterial.SelectedValue(nJOYERIA).Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.cboUnidad.SelectedValue(nJOYERIA).Text) <> Trim(Me.cboUnidad.SelectedValue(nJOYERIA).Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.cboAlmacen.SelectedValue(nJOYERIA).Text) <> Trim(Me.cboAlmacen.SelectedValue(nJOYERIA).Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.dbcProveedor.SelectedValue(nJOYERIA).Text) <> Trim(Me.dbcProveedor.SelectedValue(nJOYERIA).Tag) Then
        '            Cambios = True
        '        ElseIf Trim(Me.txtCodigodelProveedor(nJOYERIA).Text) <> Trim(Me.txtCodigodelProveedor(nJOYERIA).Tag) Then
        '            Cambios = True
        '        ElseIf _optMoneda_10.Checked <> CBool(_optMoneda_10.Tag) Then
        '            Cambios = True
        '        ElseIf Trim(txtAdicional(0).Text) <> Trim(txtAdicional(0).Tag) Then
        '            Cambios = True
        '        ElseIf Trim(txtImagen(nJOYERIA).Text) <> Trim(txtImagen(nJOYERIA).Tag) Then
        '            Cambios = True
        '        Else
        '            Cambios = False
        '        End If
        '    Case nRELOJERIA

        '        If Trim(Me.dbcMarca.Text) <> Trim(Me.dbcMarca.Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.dbcModelo.Text) <> Trim(Me.dbcModelo.Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.dbcMaterial.SelectedValue(nRELOJERIA).Text) <> Trim(Me.dbcMaterial.SelectedValue(nRELOJERIA).Tag) Then
        '            Cambios = True
        '        ElseIf Trim(cGenero) <> Trim(cGeneroTag) Then
        '            Cambios = True
        '        ElseIf Trim(cMovimiento) <> Trim(cMovimientoTag) Then
        '            Cambios = True
        '        ElseIf Trim(cCrono) <> Trim(cCronoTag) Then
        '            Cambios = True
        '        ElseIf Trim(Me.txtDescripcion(nRELOJERIA).Text) <> Trim(Me.txtDescripcion(nRELOJERIA).Tag) Then
        '            Cambios = True
        '        ElseIf Trim(cMonedaCompra) <> Trim(cMonedaCompraTag) Then
        '            Cambios = True
        '        ElseIf ModEstandar.Numerico(Me.txtPrecioenDolares(nRELOJERIA).Text) <> ModEstandar.Numerico(Me.txtPrecioenDolares(nRELOJERIA).Tag) Then
        '            Cambios = True
        '        ElseIf nCostoFactura <> nCostoFacturaTag Then
        '            Cambios = True
        '        ElseIf nCostoAdicional <> nCostoAdicionalTag Then
        '            Cambios = True
        '        ElseIf nCostoIndirectos <> nCostoIndirectosTag Then
        '            Cambios = True
        '        ElseIf nCostoReal <> nCostoRealTag Then
        '            Cambios = True

        '        ElseIf Trim(Me.cboUnidad.SelectedValue(nRELOJERIA).Text) <> Trim(Me.cboUnidad.SelectedValue(nRELOJERIA).Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.cboAlmacen.SelectedValue(nRELOJERIA).Text) <> Trim(Me.cboAlmacen.SelectedValue(nRELOJERIA).Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.dbcProveedor.SelectedValue(nRELOJERIA).Text) <> Trim(Me.dbcProveedor.SelectedValue(nRELOJERIA).Tag) Then
        '            Cambios = True
        '        ElseIf Trim(Me.txtCodigodelProveedor(nRELOJERIA).Text) <> Trim(Me.txtCodigodelProveedor(nRELOJERIA).Tag) Then
        '            Cambios = True
        '        ElseIf_optMoneda_7.Checked <> CBool(optMoneda(7).Tag) Then
        '            Cambios = True
        '        ElseIf Trim(txtImagen(nRELOJERIA).Text) <> Trim(txtImagen(nRELOJERIA).Tag) Then
        '            Cambios = True
        '        ElseIf Trim(txtAdicional(nRELOJERIA).Text) <> Trim(txtAdicional(nRELOJERIA).Tag) Then
        '            Cambios = True
        '        Else
        '            Cambios = False
        '        End If
        '    Case nVARIOS

        '        If Trim(Me.dbcFamilia.SelectedValue(1).Text) <> Trim(Me.dbcFamilia.SelectedValue(1).Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.dbcLinea.SelectedValue(1).Text) <> Trim(Me.dbcLinea.SelectedValue(1).Tag) Then
        '            Cambios = True
        '        ElseIf Trim(Me.txtDescripcion(nVARIOS).Text) <> Trim(Me.txtDescripcion(nVARIOS).Tag) Then
        '            Cambios = True
        '        ElseIf Trim(cMonedaCompra) <> Trim(cMonedaCompraTag) Then
        '            Cambios = True
        '        ElseIf ModEstandar.Numerico(Me.txtPrecioenDolares(nVARIOS).Text) <> ModEstandar.Numerico(Me.txtPrecioenDolares(nVARIOS).Tag) Then
        '            Cambios = True
        '        ElseIf nCostoFactura <> nCostoFacturaTag Then
        '            Cambios = True
        '        ElseIf nCostoAdicional <> nCostoAdicionalTag Then
        '            Cambios = True
        '        ElseIf nCostoIndirectos <> nCostoIndirectosTag Then
        '            Cambios = True
        '        ElseIf nCostoReal <> nCostoRealTag Then
        '            Cambios = True

        '        ElseIf Trim(Me.dbcMaterial.SelectedValue(nVARIOS).Text) <> Trim(Me.dbcMaterial.SelectedValue(nVARIOS).Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.cboUnidad.SelectedValue(nVARIOS).Text) <> Trim(Me.cboUnidad.SelectedValue(nVARIOS).Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.cboAlmacen.SelectedValue(nVARIOS).Text) <> Trim(Me.cboAlmacen.SelectedValue(nVARIOS).Tag) Then
        '            Cambios = True

        '        ElseIf Trim(Me.dbcProveedor.SelectedValue(nVARIOS).Text) <> Trim(Me.dbcProveedor.SelectedValue(nVARIOS).Tag) Then
        '            Cambios = True
        '        ElseIf Trim(Me.txtCodigodelProveedor(nVARIOS).Text) <> Trim(Me.txtCodigodelProveedor(nVARIOS).Tag) Then
        '            Cambios = True
        '        ElseIf _optMoneda_8.Checked <> CBool(_optMoneda_8.Tag) Then
        '            Cambios = True
        '        ElseIf Trim(txtImagen(nVARIOS).Text) <> Trim(txtImagen(nVARIOS).Tag) Then
        '            Cambios = True
        '        ElseIf Trim(txtAdicional(nVARIOS).Text) <> Trim(txtAdicional(nVARIOS).Tag) Then
        '            Cambios = True
        '        Else
        '            Cambios = False
        '        End If
        'End Select
        If Me.dbcOrigen.Text <> dbcOrigen.Tag Then Cambios = True
        If Me.txtCodArtAnterior.Text <> Me.txtCodArtAnterior.Tag Then Cambios = True
        If chkCodigoAnterior.CheckState <> CDbl(chkCodigoAnterior.Tag) Then Cambios = True

    End Function

    Public Sub Limpiar()
        On Error Resume Next
        'Validar si hubo cambios que desee Guardar
        If Cambios() And Not mblnNuevo Then
            Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Guardar el Registro
                    If Not Guardar() Then
                        Exit Sub
                    End If
                Case MsgBoxResult.No 'No hace nada y permite que se limpie la pantalla
                Case MsgBoxResult.Cancel 'Cancela la acción de limpiar la pantalla
                    Exit Sub
            End Select
        End If

        Nuevo()

        mblnNuevo = True
        mblnCambiosEnCodigo = False

        If UCase(Me.ActiveControl.Name) <> "SSTARTICULO" Then txtCodArticulo.Focus()
    End Sub

    Public Sub Nuevo()

        txtCodArticulo.Text = ""
        txtCodArticulo.Tag = ""
        nCostoFactura = 0
        nCostoFacturaTag = 0
        nCostoAdicional = 0
        nCostoAdicionalTag = 0
        nCostoIndirectos = 0
        nCostoIndirectosTag = 0
        nCostoReal = 0
        nCostoRealTag = 0
        txtCodArtAnterior.Text = ""
        txtCodArtAnterior.Tag = ""
        dbcOrigen.Text = ""
        dbcOrigen.Tag = ""
        chkCodigoAnterior.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkCodigoAnterior.Tag = System.Windows.Forms.CheckState.Unchecked

        'ModCorporativo.BuscaImagen("", Image1)
        'ModCorporativo.BuscaImagen("", Image2)
        'ModCorporativo.BuscaImagen("", Image3)

        'txtCostoFactura(sstArticulo.SelectedIndex).Text = Cifrar("0")
        'txtCostoFactura(sstArticulo.SelectedIndex).Tag = Cifrar("0")
        'txtCostoAdicional(sstArticulo.SelectedIndex).Text = Cifrar("0")
        'txtCostoAdicional(sstArticulo.SelectedIndex).Tag = Cifrar("0")
        'txtCostoIndirecto(sstArticulo.SelectedIndex).Text = Cifrar("0")
        'txtCostoIndirecto(sstArticulo.SelectedIndex).Tag = Cifrar("0")
        'txtCostoReal(sstArticulo.SelectedIndex).Text = Cifrar("0")
        'txtCostoReal(sstArticulo.SelectedIndex).Tag = Cifrar("0")

        _txtCostoFactura_0.Text = Cifrar("0")
        _txtCostoFactura_0.Tag = Cifrar("0")
        _txtCostoAdicional_0.Text = Cifrar("0")
        _txtCostoAdicional_0.Tag = Cifrar("0")
        _txtCostoIndirecto_0.Text = Cifrar("0")
        _txtCostoIndirecto_0.Tag = Cifrar("0")
        _txtCostoReal_0.Text = Cifrar("0")
        _txtCostoReal_0.Tag = Cifrar("0")

        _txtCostoFactura_1.Text = Cifrar("0")
        _txtCostoFactura_1.Tag = Cifrar("0")
        _txtCostoAdicional_1.Text = Cifrar("0")
        _txtCostoAdicional_1.Tag = Cifrar("0")
        _txtCostoIndirecto_1.Text = Cifrar("0")
        _txtCostoIndirecto_1.Tag = Cifrar("0")
        _txtCostoReal_1.Text = Cifrar("0")
        _txtCostoReal_1.Tag = Cifrar("0")

        _txtCostoFactura_2.Text = Cifrar("0")
        _txtCostoFactura_2.Tag = Cifrar("0")
        _txtCostoAdicional_2.Text = Cifrar("0")
        _txtCostoAdicional_2.Tag = Cifrar("0")
        _txtCostoIndirecto_2.Text = Cifrar("0")
        _txtCostoIndirecto_2.Tag = Cifrar("0")
        _txtCostoReal_2.Text = Cifrar("0")
        _txtCostoReal_2.Tag = Cifrar("0")





        cFamilia(0) = ""
        cFamilia(1) = ""
        cLinea(0) = ""
        cLinea(1) = ""
        cSubLinea = ""
        cSubLineaDescCorta = ""
        cTipoMaterial = ""
        cTipoMaterialDescCorta = ""
        mintJMaterial = 0
        mintRMaterial = 0
        mintVMaterial = 0
        mstrRuta = ""
        mstrArchivo = ""

        Select Case sstArticulo.SelectedIndex
            Case nJOYERIA
                mblnFueraChange = True
                mintCodArticulo = 0
                If Not mblnCambiosEnCodigo Then
                    txtDescArticulo.Text = ""
                    txtDescArticulo.Tag = txtDescArticulo.Text
                End If
                'ModCorporativo.BuscaImagen("", Image1)
                mintJFam = 0

                'dbcFamilia.SelectedValue(0).Text = cINDEFINIDA 
                'dbcFamilia.SelectedValue(0).Tag = dbcFamilia.SelectedValue(0).Text

                _dbcFamilia_0.Text = cINDEFINIDA
                _dbcFamilia_0.Tag = _dbcFamilia_0.Text

                mintJLin = 0

                'dbcLinea.SelectedValue(0).Text = cINDEFINIDA 
                'dbcLinea.SelectedValue(0).Tag = dbcLinea.SelectedValue(0).Text

                _dbcLinea_0.Text = cINDEFINIDA
                _dbcLinea_0.Tag = _dbcLinea_0.Text


                mintJSub = 0

                dbcSubLinea.Text = cINDEFINIDA
                dbcSubLinea.Tag = dbcSubLinea.Text
                mintCodKilates = 0

                dbcKilates.Text = cINDEFINIDO

                dbcKilates.Tag = dbcKilates.Text
                mblnFueraChange = False
                _txtAdicional_0.Text = ""
                _txtAdicional_0.Tag = ""

                _txtDescripcion_0.Text = ""
                _txtDescripcion_0.Tag = ""
                _lblDescripcion_0.Text = ""
                _lblDescripcion_0.Tag = _lblDescripcion_0.Tag

                ''' 27OCT2010 - MAVF
                txtMDSPeso.Text = ""
                txtMDSColor.Text = ""
                txtMDSPureza.Text = ""
                txtMDSCertificado.Text = ""
                lblEstatus.Text = ""
                ''' ************************

                'Moneda Compra
                _optMoneda_0.Checked = False
                _optMoneda_1.Checked = False
                _optMoneda_10.Checked = False
                _optMoneda_11.Checked = False
                '''cMonedaCompra = C_DOLAR
                '''cMonedaCompraTag = C_DOLAR
                cMonedaCompra = ""
                cMonedaCompraTag = ""

                'txtPrecioenDolares(nJOYERIA).Text = "0"
                'txtPrecioenDolares(nJOYERIA).Tag = "0"

                _txtPrecioenDolares_0.Text = "0"
                _txtPrecioenDolares_0.Tag = "0"



                mblnFueraChange = True
                mintJMaterial = 0

                'dbcMaterial.SelectedValue(nJOYERIA).Text = cINDEFINIDO 
                'dbcMaterial.SelectedValue(nJOYERIA).Tag = dbcMaterial.SelectedValue(nJOYERIA).Text

                _dbcMaterial_0.Text = cINDEFINIDO
                _dbcMaterial_0.Tag = _dbcMaterial_0.Text

                mintJOrigen = 0
                mintJUnidad = 0

                'cboUnidad.SelectedValue(nJOYERIA).Text = cINDEFINIDA 
                'cboUnidad.SelectedValue(nJOYERIA).Tag = cboUnidad.SelectedValue(nJOYERIA).Text 

                _cboUnidad_0.Text = cINDEFINIDA
                _cboUnidad_0.Tag = _cboUnidad_0.Text


                'cboAlmacen.SelectedValue(nJOYERIA).Text = cINDEFINIDO 
                'cboAlmacen.SelectedValue(nJOYERIA).Tag = cboAlmacen.SelectedValue(nJOYERIA).Text

                _cboAlmacen_0.Text = cINDEFINIDO
                _cboAlmacen_0.Tag = _cboAlmacen_0.Text


                mintJProv = 0

                'dbcProveedor.SelectedValue(nJOYERIA).Text = cINDEFINIDO
                'dbcProveedor.SelectedValue(nJOYERIA).Tag = cINDEFINIDO

                _dbcProveedor_0.Text = cINDEFINIDO
                _dbcProveedor_0.Tag = cINDEFINIDO


                mblnFueraChange = False

                'txtCodigodelProveedor(nJOYERIA).Text = ""
                'txtCodigodelProveedor(nJOYERIA).Tag = ""
                'txtCostoFactura(sstArticulo.SelectedIndex).ReadOnly = False
                'txtCostoIndirecto(sstArticulo.SelectedIndex).ReadOnly = False
                'txtCostoAdicional(sstArticulo.SelectedIndex).ReadOnly = False
                'txtImagen(nJOYERIA).Text = ""
                'txtImagen(nJOYERIA).Tag = ""

                _txtCodigodelProveedor_0.Text = ""
                _txtCodigodelProveedor_0.Tag = ""
                _txtCostoFactura_0.ReadOnly = False
                _txtCostoIndirecto_0.ReadOnly = False
                _txtCostoAdicional_0.ReadOnly = False
                _txtImagen_0.Text = ""
                _txtImagen_0.Tag = ""


            Case nRELOJERIA
                mblnFueraChange = True
                mintCodArticulo = 0
                'ModCorporativo.BuscaImagen("", Image2)
                txtDescArticulo.Text = ""
                txtDescArticulo.Tag = txtDescArticulo.Text
                mintRMar = 0

                dbcMarca.Text = cINDEFINIDA
                dbcMarca.Tag = cINDEFINIDA
                mintRMod = 0

                dbcModelo.Text = cINDEFINIDO
                dbcModelo.Tag = cINDEFINIDO
                mintRMaterial = 0

                'dbcMaterial.SelectedValue(nRELOJERIA).Text = cINDEFINIDO 
                'dbcMaterial.SelectedValue(nRELOJERIA).Tag = dbcMaterial.SelectedValue(nRELOJERIA).Text


                _dbcMaterial_1.Text = cINDEFINIDO
                _dbcMaterial_1.Tag = _dbcMaterial_1.Text


                _txtAdicional_1.Text = ""
                _txtAdicional_1.Tag = ""
                mblnFueraChange = False

                'Género
                _optGenero_0.Checked = False
                _optGenero_1.Checked = False
                _optGenero_2.Checked = False
                '''cGenero = "H"
                '''cGeneroTag = "H"
                cGenero = ""
                cGeneroTag = ""

                'Movimiento
                _optMovimiento_0.Checked = False
                _optMovimiento_1.Checked = False
                _optMovimiento_2.Checked = False
                '''cMovimiento = "Q"
                '''cMovimientoTag = "Q"
                cMovimiento = ""
                cMovimientoTag = ""

                'Cronómetro
                chkCrono.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkCrono.Tag = System.Windows.Forms.CheckState.Unchecked
                lCrono = False
                cCrono = ""
                cCronoTag = ""

                'Moneda Compra
                _optMoneda_2.Checked = False
                _optMoneda_3.Checked = False
                _optMoneda_7.Checked = False
                _optMoneda_6.Checked = False
                '''cMonedaCompra = C_DOLAR
                '''cMonedaCompraTag = C_DOLAR
                cMonedaCompra = ""
                cMonedaCompraTag = ""

                _txtDescripcion_1.Text = ""
                _txtDescripcion_1.Tag = ""
                _lblDescripcion_1.Text = ""
                _lblDescripcion_1.Tag = _lblDescripcion_1.Tag

                'txtPrecioenDolares(nRELOJERIA).Text = "0"
                'txtPrecioenDolares(nRELOJERIA).Tag = "0"

                _txtPrecioenDolares_1.Text = "0"
                _txtPrecioenDolares_1.Tag = "0"

                mblnFueraChange = True
                mintRUnidad = 0

                'cboUnidad.SelectedValue(nRELOJERIA).Text = cINDEFINIDA 
                'cboUnidad.SelectedValue(nRELOJERIA).Tag = cboUnidad.SelectedValue(nRELOJERIA).Text

                _cboUnidad_1.Text = cINDEFINIDA
                _cboUnidad_1.Tag = _cboUnidad_1.Text

                mintROrigen = 0

                'cboAlmacen.SelectedValue(nRELOJERIA).Text = cINDEFINIDO 
                ' cboAlmacen.SelectedValue(nRELOJERIA).Tag = cboAlmacen.SelectedValue(nRELOJERIA).Text

                _cboAlmacen_1.Text = cINDEFINIDO
                _cboAlmacen_1.Tag = _cboAlmacen_1.Text


                mintRProv = 0

                'dbcProveedor.SelectedValue(nRELOJERIA).Text = cINDEFINIDO 
                'dbcProveedor.SelectedValue(nRELOJERIA).Tag = dbcProveedor.SelectedValue(nRELOJERIA).Text

                _dbcProveedor_1.Text = cINDEFINIDO
                _dbcProveedor_1.Tag = _dbcProveedor_1.Text

                mblnFueraChange = False


                'txtCodigodelProveedor(nRELOJERIA).Text = ""
                'txtCodigodelProveedor(nRELOJERIA).Tag = ""
                'txtCostoFactura(sstArticulo.SelectedIndex).ReadOnly = False
                'txtCostoIndirecto(sstArticulo.SelectedIndex).ReadOnly = False
                'txtCostoAdicional(sstArticulo.SelectedIndex).ReadOnly = False
                'txtImagen(nRELOJERIA).Text = ""
                'txtImagen(nRELOJERIA).Tag = ""


                _txtCodigodelProveedor_1.Text = ""
                _txtCodigodelProveedor_1.Tag = ""
                _txtCostoFactura_1.ReadOnly = False
                _txtCostoIndirecto_1.ReadOnly = False
                _txtCostoAdicional_1.ReadOnly = False
                _txtImagen_1.Text = ""
                _txtImagen_1.Tag = ""

            Case nVARIOS
                mblnFueraChange = True
                mintCodArticulo = 0
                'ModCorporativo.BuscaImagen("", Image3)
                txtDescArticulo.Text = ""
                txtDescArticulo.Tag = txtDescArticulo.Text
                mintVFam = 0

                'dbcFamilia.SelectedValue(1).Text = cINDEFINIDA 
                'dbcFamilia.SelectedValue(1).Tag = dbcFamilia.SelectedValue(1).Text

                _dbcFamilia_1.Text = cINDEFINIDA
                _dbcFamilia_1.Tag = _dbcFamilia_1.Text

                mintVLin = 0

                'dbcLinea.SelectedValue(1).Text = cINDEFINIDA 
                'dbcLinea.SelectedValue(1).Tag = dbcLinea.SelectedValue(1).Text

                _dbcLinea_1.Text = cINDEFINIDA
                _dbcLinea_1.Tag = _dbcLinea_1.Text


                _txtAdicional_2.Text = ""
                _txtAdicional_2.Tag = ""

                'Moneda Compra
                _optMoneda_4.Checked = False
                _optMoneda_5.Checked = False
                _optMoneda_8.Checked = False
                _optMoneda_9.Checked = False
                '''cMonedaCompra = C_DOLAR
                '''cMonedaCompraTag = C_DOLAR
                cMonedaCompra = ""
                cMonedaCompraTag = ""

                mblnFueraChange = False

                _txtDescripcion_2.Text = ""
                _txtDescripcion_2.Tag = ""
                _lblDescripcion_2.Text = ""
                _lblDescripcion_2.Tag = _lblDescripcion_2.Tag
                'txtPrecioenDolares(nVARIOS).Text = "0"
                'txtPrecioenDolares(nVARIOS).Tag = "0"

                _txtPrecioenDolares_2.Text = "0"
                _txtPrecioenDolares_2.Tag = "0"

                mblnFueraChange = True
                mintVMaterial = 0

                'dbcMaterial.SelectedValue(nVARIOS).Text = cINDEFINIDO 
                'dbcMaterial.SelectedValue(nVARIOS).Tag = dbcMaterial.SelectedValue(nVARIOS).Text

                _dbcMaterial_2.Text = cINDEFINIDO
                _dbcMaterial_2.Tag = _dbcMaterial_2.Text

                mintVUnidad = 0

                'cboUnidad.SelectedValue(nVARIOS).Text = cINDEFINIDA 
                'cboUnidad.SelectedValue(nVARIOS).Tag = cboUnidad.SelectedValue(nVARIOS).Text


                _cboUnidad_2.Text = cINDEFINIDO
                _cboUnidad_2.Tag = _cboUnidad_2.Text

                mintVOrigen = 0

                'cboAlmacen.SelectedValue(nVARIOS).Text = cINDEFINIDO 
                'cboAlmacen.SelectedValue(nVARIOS).Tag = cboAlmacen.SelectedValue(nVARIOS).Text

                _cboAlmacen_2.Text = cINDEFINIDO
                _cboAlmacen_2.Tag = _cboAlmacen_2.Text


                mintVProv = 0

                'dbcProveedor.SelectedValue(nVARIOS).Text = cINDEFINIDO 
                'dbcProveedor.SelectedValue(nVARIOS).Tag = dbcProveedor.SelectedValue(nVARIOS).Text

                _dbcProveedor_2.Text = cINDEFINIDO
                _dbcProveedor_2.Tag = _dbcProveedor_2.Text

                mblnFueraChange = False

                'txtCodigodelProveedor(nVARIOS).Text = ""
                'txtCodigodelProveedor(nVARIOS).Tag = ""
                'txtCostoFactura(sstArticulo.SelectedIndex).ReadOnly = False
                'txtCostoIndirecto(sstArticulo.SelectedIndex).ReadOnly = False
                'txtCostoAdicional(sstArticulo.SelectedIndex).ReadOnly = False
                'txtImagen(nVARIOS).Text = ""
                'txtImagen(nVARIOS).Tag = ""

                _txtCodigodelProveedor_2.Text = ""
                _txtCodigodelProveedor_2.Tag = ""
                _txtCostoFactura_2.ReadOnly = False
                _txtCostoIndirecto_2.ReadOnly = False
                _txtCostoAdicional_2.ReadOnly = False
                _txtImagen_2.Text = ""
                _txtImagen_2.Tag = ""

        End Select
    End Sub

    Public Sub LlenaDatos()
        On Error GoTo Merr
        Dim Control As System.Windows.Forms.Control
        Dim lImporteDol As Decimal

        If Not BuscaArticulo(CInt(Numerico((txtCodArticulo.Text)))) Then
            If txtCodArticulo.Text <> "" Then
                txtCodArticulo.Text = ""
                MsgBox("El código de artículo no existe" & vbNewLine & "Verifique por favor", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            End If
            mblnLlenoDatos = False
            Limpiar()
            Exit Sub
        End If
        gStrSql = "SELECT * FROM CatArticulos WHERE CodArticulo = " & Numerico((txtCodArticulo.Text))
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute

        If Not RsGral.EOF Then
            Select Case RsGral.Fields("CodGrupo").Value
                Case gCODJOYERIA
                    sstArticulo.SelectedIndex = nJOYERIA
                    'ModCorporativo.BuscaImagen(Trim(Str(CInt(RsGral.Fields("CodArticulo").Value))), Image1)
                Case gCODRELOJERIA
                    sstArticulo.SelectedIndex = nRELOJERIA
                    'ModCorporativo.BuscaImagen(Trim(Str(CInt(RsGral.Fields("CodArticulo").Value))), Image2)
                Case gCODVARIOS
                    sstArticulo.SelectedIndex = nVARIOS
                    'ModCorporativo.BuscaImagen(Trim(Str(CInt(RsGral.Fields("CodArticulo").Value))), Image3)
            End Select

            txtCodArticulo.Text = RsGral.Fields("CodArticulo").Value
            mintCodArticulo = RsGral.Fields("CodArticulo").Value
            txtDescArticulo.Text = Trim(RsGral.Fields("DescArticulo").Value)
            txtDescArticulo.Tag = txtDescArticulo.Text

            ''' 27OCT2010 - MAVF
            txtMDSPeso.Text = Format(RsGral.Fields("mdsPeso").Value, "##0.00")
            txtMDSPeso.Tag = txtMDSPeso.Text
            txtMDSColor.Text = Trim(RsGral.Fields("mdsColor").Value)
            txtMDSColor.Tag = txtMDSColor.Text
            txtMDSPureza.Text = Trim(RsGral.Fields("mdsPureza").Value)
            txtMDSPureza.Tag = txtMDSPureza.Text
            txtMDSCertificado.Text = Trim(RsGral.Fields("mdsCertificado").Value)
            txtMDSCertificado.Tag = txtMDSCertificado.Text
            ''' ***************************************************

            nCostoFactura = CDec(Format(RsGral.Fields("CostoFactura").Value, "###,###,##0.00"))
            nCostoFacturaTag = nCostoFactura
            nCostoAdicional = CDec(Format(RsGral.Fields("CostoAdicional").Value, "###,###,##0.00"))
            nCostoAdicionalTag = nCostoAdicional
            nCostoIndirectos = CDec(Format(RsGral.Fields("CostoIndirecto").Value, "###,###,##0.00"))
            nCostoIndirectosTag = nCostoIndirectos
            nCostoReal = CDec(Format(RsGral.Fields("CostoReal").Value, "###,###,##0.00"))
            nCostoRealTag = nCostoReal
            ''Obtener los Costos en pesos
            nCostoAdicionalPesos = RsGral.Fields("CostoAdicional").Value * gcurCorpoTIPOCAMBIODOLAR
            nCostoFacturaPesos = RsGral.Fields("CostoFactura").Value * gcurCorpoTIPOCAMBIODOLAR
            nCostoIndirectosPesos = RsGral.Fields("CostoIndirecto").Value * gcurCorpoTIPOCAMBIODOLAR

            'txtCostoFactura(sstArticulo.SelectedIndex).Text = Cifrar(CStr(Format(nCostoFactura, "###,##0")))
            'txtCostoFactura(sstArticulo.SelectedIndex).Tag = txtCostoFactura(sstArticulo.SelectedIndex).Text
            'txtCostoAdicional(sstArticulo.SelectedIndex).Text = Cifrar(CStr(Format(nCostoAdicional, "###,##0")))
            'txtCostoAdicional(sstArticulo.SelectedIndex).Tag = txtCostoAdicional(sstArticulo.SelectedIndex).Text
            'txtCostoIndirecto(sstArticulo.SelectedIndex).Text = Cifrar(CStr(Format(nCostoIndirectos, "###,##0")))
            'txtCostoIndirecto(sstArticulo.SelectedIndex).Tag = txtCostoIndirecto(sstArticulo.SelectedIndex).Text
            'txtCostoReal(sstArticulo.SelectedIndex).Text = Cifrar(CStr(Format(nCostoReal, "###,##0")))
            'txtCostoReal(sstArticulo.SelectedIndex).Tag = txtCostoReal(sstArticulo.SelectedIndex).Text

            _txtCostoFactura_0.Text = Cifrar(CStr(Format(nCostoFactura, ",0")))
            _txtCostoFactura_0.Tag = _txtCostoFactura_0.Text
            _txtCostoAdicional_0.Text = Cifrar(CStr(Format(nCostoAdicional, ",0")))
            _txtCostoAdicional_0.Tag = _txtCostoAdicional_0.Text
            _txtCostoIndirecto_0.Text = Cifrar(CStr(Format(nCostoIndirectos, ",0")))
            _txtCostoIndirecto_0.Tag = _txtCostoIndirecto_0.Text
            _txtCostoReal_0.Text = Cifrar(CStr(Format(String.Concat(nCostoReal, ",0"))))
            _txtCostoReal_0.Tag = _txtCostoReal_0.Text

            _txtCostoFactura_1.Text = Cifrar(CStr(Format(nCostoFactura, ",0")))
            _txtCostoFactura_1.Tag = _txtCostoFactura_1.Text
            _txtCostoAdicional_1.Text = Cifrar(CStr(Format(nCostoAdicional, ",0")))
            _txtCostoAdicional_1.Tag = _txtCostoAdicional_1.Text
            _txtCostoIndirecto_1.Text = Cifrar(CStr(Format(nCostoIndirectos, ",0")))
            _txtCostoIndirecto_1.Tag = _txtCostoIndirecto_1.Text
            _txtCostoReal_1.Text = Cifrar(CStr(Format(String.Concat(nCostoReal, ",0"))))
            _txtCostoReal_1.Tag = _txtCostoReal_1.Text


            _txtCostoFactura_2.Text = Cifrar(CStr(Format(nCostoFactura, ",0")))
            _txtCostoFactura_2.Tag = _txtCostoFactura_2.Text
            _txtCostoAdicional_2.Text = Cifrar(CStr(Format(nCostoAdicional, ",0")))
            _txtCostoAdicional_2.Tag = _txtCostoAdicional_2.Text
            _txtCostoIndirecto_2.Text = Cifrar(CStr(Format(nCostoIndirectos, ",0")))
            _txtCostoIndirecto_2.Tag = _txtCostoIndirecto_2.Text
            _txtCostoReal_2.Text = Cifrar(CStr(Format(String.Concat(nCostoReal, ",0"))))
            _txtCostoReal_2.Tag = _txtCostoReal_2.Text


            If RsGral.Fields("CodigoAnt").Value > 0 Then
                chkCodigoAnterior.CheckState = System.Windows.Forms.CheckState.Checked
                chkCodigoAnterior.Tag = System.Windows.Forms.CheckState.Checked
                txtCodArtAnterior.Text = RsGral.Fields("CodigoAnt").Value
                txtCodArtAnterior.Tag = RsGral.Fields("CodigoAnt").Value
                dbcOrigen.Text = RsGral.Fields("OrigenAnt").Value
                dbcOrigen.Tag = RsGral.Fields("OrigenAnt").Value
                intCodAlmacenOrigen = RsGral.Fields("OrigenAnt").Value
            Else
                chkCodigoAnterior.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkCodigoAnterior.Tag = System.Windows.Forms.CheckState.Unchecked
                txtCodArtAnterior.Text = ""
                txtCodArtAnterior.Tag = ""
                dbcOrigen.Text = ""
                dbcOrigen.Tag = ""
                intCodAlmacenOrigen = RsGral.Fields("OrigenAnt").Value
            End If

            Select Case sstArticulo.SelectedIndex
                Case nJOYERIA
                    mblnFueraChange = True

                    mintJFam = IIf(IsDBNull(RsGral.Fields("CodFamilia").Value), 0, RsGral.Fields("CodFamilia").Value)

                    'dbcFamilia.SelectedValue(nJOYERIA).Text = BuscaFamilia(mintJFam)
                    'dbcFamilia.SelectedValue(nJOYERIA).Tag = dbcFamilia.SelectedValue(nJOYERIA).Text

                    _dbcFamilia_0.Text = BuscaFamilia(mintJFam)
                    _dbcFamilia_0.Tag = _dbcFamilia_0.Text

                    mintJLin = IIf(IsDBNull(RsGral.Fields("COdLinea").Value), 0, RsGral.Fields("COdLinea").Value)

                    'dbcLinea.SelectedValue(nJOYERIA).Text = BuscaLinea(mintJLin)
                    'dbcLinea.SelectedValue(nJOYERIA).Tag = dbcLinea.SelectedValue(nJOYERIA).Text

                    _dbcLinea_0.Text = BuscaLinea(mintJLin)
                    _dbcLinea_0.Tag = _dbcLinea_0.Text


                    mintJSub = IIf(IsDBNull(RsGral.Fields("CodSubLinea").Value), 0, RsGral.Fields("CodSubLinea").Value)

                    dbcSubLinea.Text = BuscaSubLinea(mintJSub)
                    dbcSubLinea.Tag = dbcSubLinea.Text



                    If Not IsDBNull(RsGral.Fields("CodKilates").Value) Then
                        mintCodKilates = RsGral.Fields("CodKilates").Value

                        dbcKilates.Text = BuscaKilataje(mintCodKilates)
                        dbcKilates.Tag = dbcKilates.Text
                    Else
                        mintCodKilates = 0

                        dbcKilates.Text = cINDEFINIDO
                        dbcKilates.Tag = dbcKilates.Text
                    End If

                    mblnFueraChange = False
                    _txtAdicional_0.Text = Trim(RsGral.Fields("Adicional").Value)
                    _txtAdicional_0.Tag = _txtAdicional_0.Text

                    'Tipo de Moneda con la que se efectuó la compra
                    _optMoneda_0.Checked = False
                    _optMoneda_1.Checked = False
                    mblnFueraChange = True

                    If Trim(RsGral.Fields("MonedaCompra").Value) = C_DOLAR Then
                        _optMoneda_0.Checked = True
                        _optMoneda_1.Checked = False
                    Else
                        _optMoneda_0.Checked = False
                        _optMoneda_1.Checked = True
                    End If
                    mblnFueraChange = False
                    cMonedaCompra = Trim(RsGral.Fields("MonedaCompra").Value)
                    cMonedaCompraTag = Trim(RsGral.Fields("MonedaCompra").Value)

                    'txtDescripcion(nJOYERIA).Text = Trim(RsGral.Fields("DescArticulo").Value)
                    'txtDescripcion(nJOYERIA).Tag = txtDescripcion(nJOYERIA).Text

                    _txtDescripcion_0.Text = Trim(RsGral.Fields("DescArticulo").Value)
                    _txtDescripcion_0.Tag = _txtDescripcion_0.Text

                    'lblDescripcion(nJOYERIA).Text = Trim(RsGral.Fields("DescArticulo").Value)
                    'lblDescripcion(nJOYERIA).Tag = lblDescripcion(nJOYERIA).Text

                    _lblDescripcion_0.Text = Trim(RsGral.Fields("DescArticulo").Value)
                    _lblDescripcion_0.Tag = _lblDescripcion_0.Text

                    'txtPrecioenDolares(nJOYERIA).Text = Format(RsGral.Fields("PrecioPubDolar").Value, "###,###,##0")
                    'txtPrecioenDolares(nJOYERIA).Tag = txtPrecioenDolares(nJOYERIA).Text

                    _txtPrecioenDolares_0.Text = Format(RsGral.Fields("PrecioPubDolar").Value, "###,###,##0")
                    _txtPrecioenDolares_0.Tag = _txtPrecioenDolares_0.Text

                    mblnFueraChange = True

                    mintJUnidad = IIf(IsDBNull(RsGral.Fields("CodUnidad").Value), 0, RsGral.Fields("CodUnidad").Value)

                    'cboUnidad.SelectedValue(nJOYERIA).Text = BuscaUnidad(mintJUnidad)
                    'cboUnidad.SelectedValue(nJOYERIA).Tag = cboUnidad.SelectedValue(nJOYERIA).Text

                    _cboUnidad_0.Text = BuscaUnidad(mintJUnidad)
                    _cboUnidad_0.Tag = _cboUnidad_0.Text

                    mintJOrigen = IIf(IsDBNull(RsGral.Fields("CodAlmacenOrigen").Value), 0, RsGral.Fields("CodAlmacenOrigen").Value)

                    'cboAlmacen.SelectedValue(nJOYERIA).Text = BuscaAlmacen(mintJOrigen)
                    'cboAlmacen.SelectedValue(nJOYERIA).Tag = cboAlmacen.SelectedValue(nJOYERIA).Text

                    _cboAlmacen_0.Text = BuscaAlmacen(mintJOrigen)
                    _cboAlmacen_0.Tag = _cboAlmacen_0.Text


                    mintJMaterial = IIf(IsDBNull(RsGral.Fields("CodTipoMaterial").Value), 0, RsGral.Fields("CodTipoMaterial").Value)

                    'dbcMaterial.SelectedValue(nJOYERIA).Text = BuscaTipoMaterial(mintJMaterial)
                    'dbcMaterial.SelectedValue(nJOYERIA).Tag = dbcMaterial.SelectedValue(nJOYERIA).Text
                    'cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintJMaterial)

                    _dbcMaterial_0.Text = BuscaTipoMaterial(mintJMaterial)
                    _dbcMaterial_0.Tag = _dbcMaterial_0.Text
                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintJMaterial)


                    mintJProv = IIf(IsDBNull(RsGral.Fields("CodProveedor").Value), 0, RsGral.Fields("CodProveedor").Value)

                    'dbcProveedor.SelectedValue(nJOYERIA).Text = BuscaProveedor(mintJProv)
                    'dbcProveedor.SelectedValue(nJOYERIA).Tag = dbcProveedor.SelectedValue(nJOYERIA).Text

                    _dbcProveedor_0.Text = BuscaProveedor(mintJProv)
                    _dbcProveedor_0.Tag = _dbcProveedor_0.Text


                    mblnFueraChange = False

                    'txtCodigodelProveedor(nJOYERIA).Text = Trim(RsGral.Fields("CodigoArticuloProv").Value)
                    'txtCodigodelProveedor(nJOYERIA).Tag = txtCodigodelProveedor(nJOYERIA).Text

                    _txtCodigodelProveedor_0.Text = Trim(RsGral.Fields("CodigoArticuloProv").Value)
                    _txtCodigodelProveedor_0.Tag = _txtCodigodelProveedor_0.Text

                    _optMoneda_10.Checked = False
                    _optMoneda_11.Checked = False
                    mblnFueraChange = True
                    If RsGral.Fields("PesosFijos").Value = True Then
                        _optMoneda_10.Checked = False
                        _optMoneda_10.Tag = False
                        _optMoneda_11.Checked = True
                        _optMoneda_11.Tag = True
                    Else
                        _optMoneda_10.Checked = True
                        _optMoneda_10.Tag = True
                        _optMoneda_11.Checked = False
                        _optMoneda_11.Tag = False
                    End If
                    mblnFueraChange = False
                    '''Bajo riesgo del cliente se habilitan los campos de costos
                    '''txtCostoFactura(sstArticulo.Tab).Locked = True
                    '''txtCostoIndirecto(sstArticulo.Tab).Locked = True
                    '''txtCostoAdicional(sstArticulo.Tab).Locked = True
                    FormaDescripcion()

                    'If Not _optMoneda_10.Checked Then lImporteDol = (CDec(Numerico(txtPrecioenDolares(0).Text)) / gcurCorpoTIPOCAMBIODOLAR) Else lImporteDol = CDec(Numerico(txtPrecioenDolares(0).Text))
                    'If lImporteDol > 0 Then
                    '    lblMargen(0).Text = Format((1 - (System.Math.Round(CDec(DesCifrar(txtCostoReal(0).Text)) / lImporteDol, 4))) * 100, "##0.00")
                    'Else
                    '    lblMargen(0).Text = "0.00"
                    'End If
                    'FormaDescripcion()

                    If Not _optMoneda_10.Checked Then lImporteDol = (CDec(Numerico(_txtPrecioenDolares_0.Text)) / gcurCorpoTIPOCAMBIODOLAR) Else lImporteDol = CDec(Numerico(_txtPrecioenDolares_0.Text))
                    If lImporteDol > 0 Then
                        _lblMargen_0.Text = Format((1 - (System.Math.Round(CDec(DesCifrar(_txtCostoReal_0.Text)) / lImporteDol, 4))) * 100, "0.00")
                    Else
                        _lblMargen_0.Text = "0.00"
                    End If
                    FormaDescripcion()

                Case nRELOJERIA
                    mblnFueraChange = True


                    mintRMar = IIf(IsDBNull(RsGral.Fields("CodMArca").Value), 0, RsGral.Fields("CodMArca").Value)

                    dbcMarca.Text = BuscaMarca(mintRMar)
                    dbcMarca.Tag = dbcMarca.Text
                    cMarca = Trim(dbcMarca.Text)


                    mintRMod = IIf(IsDBNull(RsGral.Fields("CodModelo").Value), 0, RsGral.Fields("CodModelo").Value)
                    dbcModelo.Text = BuscaModelo(mintRMod)
                    dbcModelo.Tag = Trim(dbcModelo.Text)
                    cModelo = Trim(dbcModelo.Text)


                    'mintRMaterial = IIf(IsDBNull(RsGral.Fields("CodTipoMaterial").Value), 0, RsGral.Fields("CodTipoMaterial").Value)
                    'dbcMaterial.SelectedValue(nRELOJERIA).Text = BuscaTipoMaterial(mintRMaterial)
                    'dbcMaterial.SelectedValue(nRELOJERIA).Tag = dbcMaterial.SelectedValue(nRELOJERIA).Text

                    mintRMaterial = IIf(IsDBNull(RsGral.Fields("CodTipoMaterial").Value), 0, RsGral.Fields("CodTipoMaterial").Value)
                    _dbcMaterial_1.Text = BuscaTipoMaterial(mintRMaterial)
                    _dbcMaterial_1.Tag = _dbcMaterial_1.Text

                    'txtAdicional(1).Text = Trim(RsGral.Fields("Adicional").Value)
                    'txtAdicional(1).Tag = txtAdicional(1).Text

                    _txtAdicional_1.Text = Trim(RsGral.Fields("Adicional").Value)
                    _txtAdicional_1.Tag = _txtAdicional_1.Text

                    'cTipoMaterial = Trim(dbcMaterial.SelectedValue(nRELOJERIA).Text)

                    cTipoMaterial = Trim(_dbcMaterial_1.Text)


                    mblnFueraChange = False

                    'Género
                    _optGenero_0.Checked = False
                    _optGenero_1.Checked = False
                    _optGenero_2.Checked = False
                    Select Case Trim(RsGral.Fields("Genero").Value)
                        Case "H"
                            _optGenero_0.Checked = True
                        Case "D"
                            _optGenero_1.Checked = True
                        Case "M"
                            _optGenero_2.Checked = True
                    End Select
                    cGenero = RsGral.Fields("Genero").Value
                    cGeneroTag = RsGral.Fields("Genero").Value

                    'Movimiento
                    _optMovimiento_0.Checked = False
                    _optMovimiento_1.Checked = False
                    _optMovimiento_2.Checked = False
                    Select Case Trim(RsGral.Fields("Movimiento").Value)
                        Case "Q"
                            _optMovimiento_0.Checked = True
                        Case "AUT"
                            _optMovimiento_1.Checked = True
                        Case "MAN"
                            _optMovimiento_2.Checked = True
                    End Select
                    cMovimiento = RsGral.Fields("Movimiento").Value
                    cMovimientoTag = cMovimiento

                    Select Case True
                        Case RsGral.Fields("Crono").Value
                            chkCrono.CheckState = System.Windows.Forms.CheckState.Checked
                            cCrono = C_CRONO
                            cCronoTag = cCrono
                            lCrono = True
                        Case Else
                            chkCrono.CheckState = System.Windows.Forms.CheckState.Unchecked
                            cCrono = ""
                            cCronoTag = cCrono
                            lCrono = False
                    End Select

                    'Tipo de Moneda con la que se efectuó la compra
                    _optMoneda_2.Checked = False
                    _optMoneda_3.Checked = False
                    mblnFueraChange = True
                    If Trim(RsGral.Fields("MonedaCompra").Value) = C_DOLAR Then
                        _optMoneda_2.Checked = True
                        _optMoneda_3.Checked = False
                    Else
                        _optMoneda_2.Checked = False
                        _optMoneda_3.Checked = True
                    End If
                    mblnFueraChange = False
                    cMonedaCompra = Trim(RsGral.Fields("MonedaCompra").Value)
                    cMonedaCompraTag = Trim(RsGral.Fields("MonedaCompra").Value)
                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintRMaterial)

                    'txtDescripcion(nRELOJERIA).Text = Trim(RsGral.Fields("DescArticulo").Value)
                    'txtDescripcion(nRELOJERIA).Tag = txtDescripcion(nRELOJERIA).Text
                    'lblDescripcion(nRELOJERIA).Text = Trim(RsGral.Fields("DescArticulo").Value)
                    'lblDescripcion(nRELOJERIA).Tag = lblDescripcion(nRELOJERIA).Text

                    _txtDescripcion_1.Text = Trim(RsGral.Fields("DescArticulo").Value)
                    _txtDescripcion_1.Tag = _txtDescripcion_1.Text
                    _lblDescripcion_1.Text = Trim(RsGral.Fields("DescArticulo").Value)
                    _lblDescripcion_1.Tag = _lblDescripcion_1.Text

                    _txtPrecioenDolares_1.Text = Format(RsGral.Fields("PrecioPubDolar").Value, "###,###,##0")
                    _txtPrecioenDolares_1.Tag = _txtPrecioenDolares_1.Text

                    mblnFueraChange = True


                    mintRUnidad = IIf(IsDBNull(RsGral.Fields("CodUnidad").Value), 0, RsGral.Fields("CodUnidad").Value)
                    _cboUnidad_1.Text = BuscaUnidad(mintRUnidad)
                    _cboUnidad_1.Tag = _cboUnidad_1.Text


                    mintROrigen = IIf(IsDBNull(RsGral.Fields("CodAlmacenOrigen").Value), 0, RsGral.Fields("CodAlmacenOrigen").Value)
                    _cboAlmacen_1.Text = BuscaAlmacen(mintROrigen)
                    _cboAlmacen_1.Tag = _cboAlmacen_1.Text


                    mintRProv = IIf(IsDBNull(RsGral.Fields("CodProveedor").Value), 0, RsGral.Fields("CodProveedor").Value)

                    'dbcProveedor.SelectedValue(nRELOJERIA).Text = BuscaProveedor(mintRProv)
                    'dbcProveedor.SelectedValue(nRELOJERIA).Tag = dbcProveedor.SelectedValue(nRELOJERIA).Text

                    _dbcProveedor_1.Text = BuscaProveedor(mintRProv)
                    _dbcProveedor_1.Tag = _dbcProveedor_1.Text

                    mblnFueraChange = False

                    'txtCodigodelProveedor(nRELOJERIA).Text = Trim(RsGral.Fields("CodigoArticuloProv").Value)
                    'txtCodigodelProveedor(nRELOJERIA).Tag = txtCodigodelProveedor(nRELOJERIA).Text

                    _txtCodigodelProveedor_1.Text = Trim(RsGral.Fields("CodigoArticuloProv").Value)
                    _txtCodigodelProveedor_1.Tag = _txtCodigodelProveedor_1.Text


                    _optMoneda_7.Checked = False
                    _optMoneda_6.Checked = False
                    mblnFueraChange = True
                    If RsGral.Fields("PesosFijos").Value = True Then
                        _optMoneda_7.Checked = False
                        _optMoneda_7.Tag = False
                        _optMoneda_6.Checked = True
                        _optMoneda_6.Tag = True
                    Else
                        _optMoneda_7.Checked = True
                        _optMoneda_7.Tag = True
                        _optMoneda_6.Checked = False
                        _optMoneda_6.Tag = False
                    End If
                    mblnFueraChange = False
                    '''Bajo riesgo del cliente se habilitan los campos de costos
                    '''txtCostoFactura(sstArticulo.Tab).Locked = True
                    '''txtCostoIndirecto(sstArticulo.Tab).Locked = True
                    '''txtCostoAdicional(sstArticulo.Tab).Locked = True

                    '''If Not_optMoneda_7.Value Then lImporteDol = (CCur(Numerico(txtPrecioenDolares(1).text)) / gcurCorpoTIPOCAMBIODOLAR) Else lImporteDol = CCur(Numerico(txtPrecioenDolares(1).text))
                    'If_optMoneda_7.Checked Then lImporteDol = CDec(Numerico(txtPrecioenDolares(1).Text)) Else lImporteDol = (CDec(Numerico(txtPrecioenDolares(1).Text)) / gcurCorpoTIPOCAMBIODOLAR)
                    'If lImporteDol > 0 Then
                    '    lblMargen(1).Text = VB6.Format((1 - (System.Math.Round(CDec(DesCifrar(txtCostoReal(1).Text)) / lImporteDol, 4))) * 100, "##0.00")
                    'Else
                    '    lblMargen(1).Text = "0.00"
                    'End If
                    'FormaDescripcion()

                    If Not _optMoneda_7.Checked Then lImporteDol = CDec(Numerico(_txtPrecioenDolares_1.Text)) Else lImporteDol = (Convert.ToDecimal((Numerico(_txtPrecioenDolares_1.Text)) / gcurCorpoTIPOCAMBIODOLAR))
                    If lImporteDol > 0 Then
                        _lblMargen_1.Text = Format((1 - (System.Math.Round(CDec(DesCifrar(_txtCostoReal_1.Text)) / lImporteDol, 4))) * 100, "##0.00")
                        'lblMargen(1).Caption = Format((1 - (Round(CCur(DesCifrar(txtCostoReal(1).text)) / lImporteDol, 4))) * 100, "##0.00")
                    Else
                        _lblMargen_1.Text = "0.00"
                    End If
                    FormaDescripcion()

                Case nVARIOS
                    mblnFueraChange = True


                    mintVFam = IIf(IsDBNull(RsGral.Fields("CodFamilia").Value), 0, RsGral.Fields("CodFamilia").Value)
                    _dbcFamilia_1.Text = BuscaFamilia(mintVFam)
                    _dbcFamilia_1.Tag = _dbcFamilia_1.Text


                    mintVLin = IIf(IsDBNull(RsGral.Fields("COdLinea").Value), 0, RsGral.Fields("COdLinea").Value)
                    _dbcLinea_1.Text = BuscaLinea(mintVLin)
                    _dbcLinea_1.Tag = _dbcLinea_1.Text

                    _txtAdicional_2.Text = Trim(RsGral.Fields("Adicional").Value)
                    _txtAdicional_2.Tag = _txtAdicional_2.Text
                    mblnFueraChange = False

                    'Tipo de Moneda con la que se efectuó la compra
                    _optMoneda_4.Checked = False
                    _optMoneda_5.Checked = False
                    mblnFueraChange = True
                    If Trim(RsGral.Fields("MonedaCompra").Value) = C_DOLAR Then
                        _optMoneda_4.Checked = True
                        _optMoneda_5.Checked = False
                    Else
                        _optMoneda_4.Checked = False
                        _optMoneda_5.Checked = True
                    End If
                    mblnFueraChange = False
                    cMonedaCompra = Trim(RsGral.Fields("MonedaCompra").Value)
                    cMonedaCompraTag = Trim(RsGral.Fields("MonedaCompra").Value)

                    _txtDescripcion_2.Text = Trim(RsGral.Fields("DescArticulo").Value)
                    _txtDescripcion_2.Tag = _txtDescripcion_2.Text
                    _lblDescripcion_2.Text = Trim(RsGral.Fields("DescArticulo").Value)
                    _lblDescripcion_2.Tag = _lblDescripcion_2.Text

                    'txtPrecioenDolares(nVARIOS).Text = Format(RsGral.Fields("PrecioPubDolar").Value, "###,###,##0")
                    'txtPrecioenDolares(nVARIOS).Tag = txtPrecioenDolares(nVARIOS).Text

                    _txtPrecioenDolares_2.Text = Format(RsGral.Fields("PrecioPubDolar").Value, ",0")
                    _txtPrecioenDolares_2.Tag = _txtPrecioenDolares_2.Text


                    mblnFueraChange = True

                    mintVUnidad = IIf(IsDBNull(RsGral.Fields("CodUnidad").Value), 0, RsGral.Fields("CodUnidad").Value)
                    _cboUnidad_2.Text = BuscaUnidad(mintVUnidad)
                    _cboUnidad_2.Tag = _cboUnidad_2.Text


                    mintVOrigen = IIf(IsDBNull(RsGral.Fields("CodAlmacenOrigen").Value), 0, RsGral.Fields("CodAlmacenOrigen").Value)
                    _cboAlmacen_2.Text = BuscaAlmacen(mintVOrigen)
                    _cboAlmacen_2.Tag = _cboAlmacen_2.Text


                    mintVMaterial = IIf(IsDBNull(RsGral.Fields("CodTipoMaterial").Value), 0, RsGral.Fields("CodTipoMaterial").Value)
                    _dbcMaterial_2.Text = BuscaTipoMaterial(mintVMaterial)
                    _dbcMaterial_2.Tag = _dbcMaterial_2.Text
                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintVMaterial)

                    mintVProv = IIf(IsDBNull(RsGral.Fields("CodProveedor").Value), 0, RsGral.Fields("CodProveedor").Value)

                    'dbcProveedor.SelectedValue(nVARIOS).Text = BuscaProveedor(mintVProv)
                    'dbcProveedor.SelectedValue(nVARIOS).Tag = dbcProveedor.SelectedValue(nVARIOS).Text

                    _dbcProveedor_2.Text = BuscaProveedor(mintVProv)
                    _dbcProveedor_2.Tag = _dbcProveedor_2.Text


                    mblnFueraChange = False

                    _txtCodigodelProveedor_2.Text = Trim(RsGral.Fields("CodigoArticuloProv").Value)
                    _txtCodigodelProveedor_2.Tag = _txtCodigodelProveedor_2.Text


                    _optMoneda_8.Checked = False
                    _optMoneda_9.Checked = False
                    mblnFueraChange = True
                    If RsGral.Fields("PesosFijos").Value = True Then
                        _optMoneda_8.Checked = False
                        _optMoneda_8.Tag = False
                        _optMoneda_9.Checked = True
                        _optMoneda_9.Tag = True
                    Else
                        _optMoneda_8.Checked = True
                        _optMoneda_8.Tag = True
                        _optMoneda_9.Checked = False
                        _optMoneda_9.Tag = False
                    End If
                    mblnFueraChange = False
                    '''Bajo riesgo del cliente se habilitan los campos de costos
                    '''txtCostoFactura(sstArticulo.Tab).Locked = True
                    '''txtCostoIndirecto(sstArticulo.Tab).Locked = True
                    '''txtCostoAdicional(sstArticulo.Tab).Locked = True
                    FormaDescripcion()

                    'If Not _optMoneda_8.Checked Then lImporteDol = (CDec(Numerico(txtPrecioenDolares(2).Text)) / gcurCorpoTIPOCAMBIODOLAR) Else lImporteDol = CDec(Numerico(txtPrecioenDolares(2).Text))
                    'If lImporteDol > 0 Then
                    '    lblMargen(2).Text = Format((1 - (System.Math.Round(CDec(DesCifrar(txtCostoReal(2).Text)) / lImporteDol, 4))) * 100, "##0.00")
                    'Else
                    '    lblMargen(2).Text = "0.00"
                    'End If



                    If Not _optMoneda_8.Checked Then lImporteDol = (CDec(Numerico(_txtPrecioenDolares_2.Text)) / gcurCorpoTIPOCAMBIODOLAR) Else lImporteDol = CDec(Numerico(_txtPrecioenDolares_2.Text))
                    If lImporteDol > 0 Then
                        _lblMargen_2.Text = Format((1 - (System.Math.Round(CDec(DesCifrar(_txtCostoReal_2.Text)) / lImporteDol, 4))) * 100, "0.00")
                    Else
                        _lblMargen_2.Text = "0.00"
                    End If


            End Select

        Else
            MsgBox("El código de artículo no existe " & vbNewLine & "Verique por favor", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            Limpiar()
            mblnLlenoDatos = False
        End If
        mblnLlenoDatos = True
        mblnNuevo = False
        mblnCambiosEnCodigo = False
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function BuscaArticulo(ByRef Codigo As Integer) As Boolean
        On Error GoTo Merr
        gStrSql = "SELECT codArticulo FROM CatArticulos WHERE CodArticulo = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaArticulo = True
        Else
            BuscaArticulo = False
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Function BuscaFamilia(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                gStrSql = "SELECT codFamilia, DescFamilia FROM CatFamilias WHERE CodGrupo = " & gCODJOYERIA & " AND CodFamilia = " & Codigo
            Case nVARIOS
                gStrSql = "SELECT codFamilia, DescFamilia FROM CatFamilias WHERE CodGrupo = " & gCODVARIOS & " AND CodFamilia = " & Codigo
            Case Else
                BuscaFamilia = ""
                Exit Function
        End Select
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaFamilia = Trim(rsLocal.Fields("DescFamilia").Value)
        Else
            BuscaFamilia = cINDEFINIDA
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaLinea(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                gStrSql = "SELECT codFamilia, codLinea, DescLinea FROM CatLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & mintJFam & " AND CodLinea = " & Codigo
            Case nVARIOS
                gStrSql = "SELECT codFamilia, codLinea, DescLinea FROM CatLineas WHERE CodGrupo = " & gCODVARIOS & " and CodFamilia = " & mintVFam & " AND CodLinea = " & Codigo
            Case Else
                BuscaLinea = ""
                Exit Function
        End Select
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaLinea = Trim(rsLocal.Fields("DescLinea").Value)
        Else
            BuscaLinea = cINDEFINIDA
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaSubLinea(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codFamilia, codLinea, codSubLinea, DescSubLinea FROM CatSubLineas WHERE CodGrupo = " & gCODJOYERIA & " and CodFamilia = " & mintJFam & " AND CodLinea = " & mintJLin & " AND CodSubLinea = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaSubLinea = Trim(rsLocal.Fields("DescSubLinea").Value)
        Else
            BuscaSubLinea = cINDEFINIDA
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaKilataje(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codKilates, DescKilates FROM CatKilates WHERE CodKilates = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaKilataje = Trim(rsLocal.Fields("DescKilates").Value)
        Else
            BuscaKilataje = cINDEFINIDO
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaMarca(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codMarca, DescMarca FROM CatMarcas WHERE CodGrupo = " & gCODRELOJERIA & " AND CodMarca = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaMarca = Trim(rsLocal.Fields("DescMarca").Value)
        Else
            BuscaMarca = cINDEFINIDA
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaModelo(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codModelo, DescModelo FROM CatModelos WHERE CodGrupo = " & gCODRELOJERIA & " AND CodMarca = " & mintRMar & " AND CodModelo = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaModelo = Trim(rsLocal.Fields("DescModelo").Value)
        Else
            BuscaModelo = cINDEFINIDO
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaTipoMaterial(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codTipoMaterial, DescTipoMaterial FROM CatTipoMaterial WHERE CodTipoMaterial = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaTipoMaterial = Trim(rsLocal.Fields("DescTipoMaterial").Value)
        Else
            BuscaTipoMaterial = cINDEFINIDO
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
        Return BuscaTipoMaterial
    End Function

    Public Function BuscaTipoMaterialDescCorta_Nombre(ByRef TipoMat As String) As Integer
        On Error GoTo Merr
        gStrSql = "SELECT codTipoMaterial, ltrim(rtrim(DescCorta)) as DescCorta FROM CatTipoMaterial WHERE DescTipoMaterial = '" & TipoMat & "' "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaTipoMaterialDescCorta_Nombre = CInt(Trim(rsLocal.Fields("CodTipoMaterial").Value))
        Else
            BuscaTipoMaterialDescCorta_Nombre = 0
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Function BuscaTipoMaterialDescCorta(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codTipoMaterial, DescCorta FROM CatTipoMaterial WHERE CodTipoMaterial = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaTipoMaterialDescCorta = Trim(rsLocal.Fields("DescCorta").Value)
        Else
            BuscaTipoMaterialDescCorta = ""
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Function BuscaSubLDescCorta(ByRef lGpo As Integer, ByRef lFam As Integer, ByRef lLin As Integer, ByRef lSubL As Integer) As String
        Dim rsLocal As ADODB.Recordset
        On Error GoTo Merr

        gStrSql = "SELECT codSubLinea, DescCorta FROM CatSubLineas WHERE CodGrupo = " & lGpo & " And CodFamilia = " & lFam & " And CodLinea = " & lLin & " And CodSubLinea = " & lSubL & " "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaSubLDescCorta = Trim(rsLocal.Fields("DescCorta").Value)
        Else
            BuscaSubLDescCorta = ""
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    '''06AGO2007 - MAVF
    Public Function BuscaSubLDescCortaDesc(ByRef lGpo As Integer, ByRef lFam As Integer, ByRef lLin As Integer, ByRef lSubL As String) As String
        Dim rsLocal As ADODB.Recordset
        On Error GoTo Merr

        If lSubL <> "" Then
            gStrSql = "SELECT codSubLinea, DescCorta FROM CatSubLineas WHERE CodGrupo = " & lGpo & " And CodFamilia = " & lFam & " And CodLinea = " & lLin & " And DescSubLinea Like '" & lSubL & "%' Order by DescSubLinea "
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.UP_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            rsLocal = Cmd.Execute
            If rsLocal.RecordCount > 0 Then
                BuscaSubLDescCortaDesc = Trim(rsLocal.Fields("DescCorta").Value)
            Else
                BuscaSubLDescCortaDesc = ""
            End If
        Else
            BuscaSubLDescCortaDesc = ""
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Function BuscaUnidad(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codUnidad, DescUnidad FROM CatUnidades WHERE CodUnidad = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaUnidad = Trim(rsLocal.Fields("DescUnidad").Value)
        Else
            BuscaUnidad = cINDEFINIDA
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaAlmacen(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codAlmacenOrigen, DescAlmacenOrigen FROM CatOrigen WHERE CodAlmacenOrigen = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaAlmacen = Trim(rsLocal.Fields("DescAlmacenorigen").Value)
        Else
            BuscaAlmacen = cINDEFINIDO
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Function

    Public Function BuscaProveedor(ByRef Codigo As Integer) As String
        On Error GoTo Merr
        gStrSql = "SELECT codProvAcreed, DescProvAcreed FROM CatProvAcreed WHERE codProvAcreed = " & Codigo
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        rsLocal = Cmd.Execute
        If rsLocal.RecordCount > 0 Then
            BuscaProveedor = Trim(rsLocal.Fields("DescProvACreed").Value)
        Else
            BuscaProveedor = cINDEFINIDO
        End If

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Sub FormaDescripcion()
        Dim strAdicional As String

        If Trim(cSubLineaDescCorta) = Trim(cINDEFINIDA) Then cSubLineaDescCorta = ""
        If Trim(cTipoMaterialDescCorta) = Trim(cINDEFINIDA) Then cTipoMaterialDescCorta = ""
        If Trim(cKilates) = Trim(cSINKILATES) Then cKilates = " " '''08NOV2010 - MAVF

        'cFamilia(0) = ""
        'cLinea(0) = ""
        'cSubLinea = ""
        'cKilates = ""
        'cTipoMaterial = ""
        'cMarca = ""
        'cModelo = ""
        'cFamilia(1) = ""
        'cLinea(1) = ""
        'strAdicional = ""

        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                cFamilia(0) = ""
                cLinea(0) = IIf(Trim(cLinea(0)) = Trim(cINDEFINIDA), "", Trim(cLinea(0))) & " "
                cSubLinea = IIf(Trim(cSubLinea) = Trim(cINDEFINIDA), "", Trim(cSubLineaDescCorta)) & " "
                cKilates = IIf(Trim(cKilates) = Trim(cINDEFINIDO), " ", Trim(cKilates)) & " "
                cTipoMaterial = IIf(Trim(cTipoMaterial) = Trim(cINDEFINIDO), "", Trim(cTipoMaterialDescCorta)) & " "
                strAdicional = Trim(_txtAdicional_0.Text)

                cDescripcion = cFamilia(0) & cLinea(0) & cSubLinea & cKilates & cTipoMaterial & strAdicional
            Case nRELOJERIA
                cMarca = IIf(Trim(cMarca) = Trim(cINDEFINIDA), "", Trim(cMarca) & " ")
                cModelo = IIf(Trim(cModelo) = Trim(cINDEFINIDO), "", Trim(cModelo) & " ")
                cTipoMaterial = IIf(Trim(cTipoMaterial) = Trim(cINDEFINIDO), " ", Trim(cTipoMaterialDescCorta) & " ")
                strAdicional = Trim(_txtAdicional_1.Text)

                cDescripcion = cMarca & cModelo & Trim(cGenero) & " " & Trim(cMovimiento) & " " & Trim(cCrono) & " " & cTipoMaterial & strAdicional
            Case nVARIOS
                cFamilia(1) = IIf(Trim(cFamilia(1)) = Trim(cINDEFINIDA), "", Trim(cFamilia(1)) & " ")
                cLinea(1) = IIf(Trim(cLinea(1)) = Trim(cINDEFINIDA), "", Trim(cLinea(1)) & " ")
                cTipoMaterial = IIf(Trim(cTipoMaterial) = Trim(cINDEFINIDO), "", Trim(cTipoMaterialDescCorta)) & " "
                strAdicional = Trim(_txtAdicional_2.Text)

                cDescripcion = cFamilia(1) & cLinea(1) & cTipoMaterial & strAdicional
        End Select
        _txtDescripcion_0.Text = Trim(cDescripcion)
        _txtDescripcion_0.Tag = Trim(cDescripcion)
    End Sub

    Private Sub chkCodigoAnterior_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCodigoAnterior.CheckStateChanged
        If chkCodigoAnterior.CheckState = System.Windows.Forms.CheckState.Checked Then
            txtCodArtAnterior.Enabled = True
            dbcOrigen.Enabled = True
        Else
            txtCodArtAnterior.Enabled = False
            dbcOrigen.Enabled = False
        End If
    End Sub

    Private Sub chkCrono_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCrono.CheckStateChanged
        If Me.chkCrono.CheckState = System.Windows.Forms.CheckState.Checked Then
            cCrono = C_CRONO
            lCrono = True
        Else
            cCrono = ""
            lCrono = False
        End If
        Call Me.FormaDescripcion()
    End Sub

    Private Sub cmdBuscarImagen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBuscarImagen.Click
        Dim Index As Integer = cmdBuscarImagen.GetIndex(eventSender)
        'frmCorpoBuscarImagen.ShowDialog()

        If mstrArchivo = "" Then Exit Sub
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                ModCorporativo.BuscaImagenArticulosNuevos(mstrRuta, mstrArchivo, Image1)
            Case nRELOJERIA
                ModCorporativo.BuscaImagenArticulosNuevos(mstrRuta, mstrArchivo, Image2)
            Case nVARIOS
                ModCorporativo.BuscaImagenArticulosNuevos(mstrRuta, mstrArchivo, Image3)
        End Select
    End Sub

    Private Sub dbcKilates_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcKilates.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        '''08NOV2010 - MAVF
        If Trim(cKilates) = Trim(cSINKILATES) Then
            cKilates = " "
        Else

            cKilates = Trim(Me.dbcKilates.Text)
        End If
        ''' ***********************************
        Call FormaDescripcion()

        If mblnFueraChange Then
            Exit Sub
        End If


        lStrSql = "SELECT codKilates, LTrim(RTrim(descKilates)) as descKilates FROM CatKilates Where LTrim(RTrim(descKilates)) LIKE '" & Trim(Me.dbcKilates.Text) & "%' Order by DescKilates "
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcKilates))


        If Trim(Me.dbcKilates.Text) = "" Then
            mintCodKilates = 0
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcKilates_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcKilates.Enter
        Pon_Tool()
        gStrSql = "SELECT codKilates, LTrim(RTrim(descKilates)) as descKilates FROM CatKilates Order by DescKilates "
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcKilates))
    End Sub

    Private Sub dbcKilates_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcKilates.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcSubLinea.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcKilates_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcKilates.KeyUp
        Dim Aux As String

        Aux = Trim(Me.dbcKilates.Text)
        'If Me.dbcKilates.SelectedItem <> 0 Then
        'dbcKilates_Leave(dbcKilates, New System.EventArgs())
        'End If

        Me.dbcKilates.Text = Aux
    End Sub

    Private Sub dbcKilates_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcKilates.Leave
        Dim Aux As Integer
        Dim cDescripcion As String

        cDescripcion = Trim(Me.dbcKilates.Text)
        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "%'")
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        gStrSql = "SELECT codKilates, LTrim(RTrim(descKilates)) as descKilates FROM CatKilates Where LTrim(RTrim(descKilates)) = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
        Aux = mintCodKilates
        mintCodKilates = 0
        ModDCombo.DCLostFocus((Me.dbcKilates), gStrSql, mintCodKilates)

        '''08NOV2010 - MAVF
        If Trim(cKilates) = Trim(cSINKILATES) Then
            cKilates = " " '''No agrega nada
        Else

            cKilates = Trim(Me.dbcKilates.Text) '''Pone la descripción del kilataje
        End If
        ''' ************************************

        If Aux <> mintCodKilates Then
            If mintCodKilates = 0 Then
                mblnFueraChange = True

                Me.dbcKilates.Text = cINDEFINIDO
                mblnFueraChange = False
            End If
        Else
            If mintCodKilates = 0 Then
                mblnFueraChange = True

                Me.dbcKilates.Text = cINDEFINIDO
                mblnFueraChange = False
            End If
        End If
        Call FormaDescripcion()
    End Sub

    Private Sub dbcKilates_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcKilates.MouseUp
        Dim Aux As String

        Aux = Trim(Me.dbcKilates.Text)
        'If Me.dbcKilates.SelectedItem <> 0 Then
        '    dbcKilates_Leave(dbcKilates, New System.EventArgs())
        'End If

        Me.dbcKilates.Text = Aux
    End Sub

    Private Sub dbcMarca_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcMarca.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String


        cMarca = Trim(Me.dbcMarca.Text)
        Call FormaDescripcion()

        If mblnFueraChange Then Exit Sub


        lStrSql = "SELECT codMarca, LTrim(RTrim(descMarca)) as descMarca FROM catMarcas Where codGrupo = " & gCODRELOJERIA & " and descMarca LIKE '" & Trim(Me.dbcMarca.Text) & "%' Order by DescMarca "
        ModDCombo.DCChange(lStrSql, tecla, dbcMarca)


        If Trim(Me.dbcMarca.Text) = "" Then
            mblnFueraChange = True
            mintRMar = 0
            mintRMod = 0

            Me.dbcModelo.Text = cINDEFINIDO
            mblnFueraChange = False
        End If

MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcMarca_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMarca.Enter
        Pon_Tool()
        gStrSql = "SELECT codMarca, LTrim(RTrim(descMarca)) as descMarca FROM catMarcas Where codGrupo = " & gCODRELOJERIA & " ORDER BY descMarca "
        ModDCombo.DCGotFocus(gStrSql, dbcMarca)
    End Sub

    Private Sub dbcMarca_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcMarca.KeyDown
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Me.sstArticulo.Focus()
            Case Else
                tecla = eventArgs.KeyCode
        End Select
    End Sub

    Private Sub dbcMarca_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMarca.Leave
        Dim I As Integer
        Dim Aux As Integer
        Dim cDescripcion As String

        cDescripcion = Trim(Me.dbcMarca.Text)

        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        gStrSql = "SELECT codMarca, LTrim(RTrim(descMarca)) as descMarca FROM catMarcas Where codGrupo = " & gCODRELOJERIA & " and descMarca LIKE '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
        Aux = mintRMar
        mintRMar = 0
        ModDCombo.DCLostFocus(dbcMarca, gStrSql, mintRMar)

        cMarca = Trim(Me.dbcMarca.Text)
        If Aux <> mintRMar Then
            mblnFueraChange = True
            If mintRMar = 0 Then

                Me.dbcMarca.Text = cINDEFINIDA
            End If

            Me.dbcModelo.Text = cINDEFINIDA
            mblnFueraChange = False
        Else
            If mintRMar = 0 Then
                mblnFueraChange = True

                Me.dbcMarca.Text = cINDEFINIDA
                mblnFueraChange = False
            End If
        End If
        Call FormaDescripcion()
    End Sub

    Private Sub dbcMarca_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcMarca.MouseUp
        Dim Aux As String

        Aux = Trim(Me.dbcMarca.Text)
        'If Me.dbcMarca.SelectedItem <> 0 Then
        '    dbcMarca_Leave(dbcMarca, New System.EventArgs())
        'End If

        Me.dbcMarca.Text = Aux
    End Sub

    Private Sub dbcModelo_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcModelo.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        cModelo = Trim(dbcModelo.Text)
        FormaDescripcion()

        If mblnFueraChange Then Exit Sub


        lStrSql = "SELECT codModelo, LTrim(RTrim(descModelo)) as descModelo FROM catModelos Where codGrupo = " & gCODRELOJERIA & " and codMarca = " & mintRMar & " and descModelo LIKE '" & Trim(dbcModelo.Text) & "%' Order by DescModelo "
        ModDCombo.DCChange(lStrSql, tecla, dbcModelo)


        If Trim(dbcModelo.Text) = "" Then mintRMod = 0

MError:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub dbcModelo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcModelo.Enter
        Pon_Tool()
        gStrSql = "SELECT codModelo, LTrim(RTrim(descModelo)) as descModelo FROM catModelos Where codGrupo = " & gCODRELOJERIA & " and codMarca = " & mintRMar & " ORDER BY descModelo "
        ModDCombo.DCGotFocus(gStrSql, dbcModelo)
    End Sub

    Private Sub dbcModelo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcModelo.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            dbcMarca.Focus()
            eventSender.KeyCode = 0
        End If
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Return Then

            If Not _optGenero_0.Checked And Not _optGenero_1.Checked And Not _optGenero_2.Checked Then
                _optGenero_0.Focus()
            ElseIf _optGenero_0.Checked Then
                _optGenero_0.Focus()
            ElseIf _optGenero_1.Checked Then
                _optGenero_1.Focus()
            ElseIf _optGenero_2.Checked Then
                _optGenero_2.Focus()
            End If

        End If

        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcModelo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcModelo.Leave
        Dim I As Integer
        Dim Aux As Integer
        Dim cDescripcion As String

        If mblnFueraChange Then Exit Sub

        cDescripcion = Trim(Me.dbcModelo.Text)
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        gStrSql = "SELECT codModelo, LTrim(RTrim(descModelo)) as descModelo FROM catModelos Where codGrupo = " & gCODRELOJERIA & " and codMarca = " & mintRMar & " and descModelo = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
        Aux = mintRMod
        mintRMod = 0
        ModDCombo.DCLostFocus(dbcModelo, gStrSql, mintRMod)

        cModelo = Trim(Me.dbcModelo.Text)
        If mintRMod = 0 Then
            mblnFueraChange = True

            Me.dbcModelo.Text = cINDEFINIDO
            mblnFueraChange = False
        End If
        FormaDescripcion()

    End Sub

    Private Sub dbcModelo_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcModelo.MouseUp
        Dim Aux As String

        Aux = Trim(Me.dbcModelo.Text)
        'If Me.dbcModelo.SelectedItem <> 0 Then
        '    dbcModelo_Leave(dbcModelo, New System.EventArgs())
        'End If

        Me.dbcModelo.Text = Aux
    End Sub

    Private Sub dbcOrigen_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcOrigen.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcOrigen" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacenOrigen, CodAlmacenOrigen AS DescAlmacen From CatOrigen WHERE DescAlmacenOrigen LIKE '" & Trim(dbcOrigen.Text) & "%' ORDER BY DescAlmacenOrigen "
        DCChange(gStrSql, tecla)
        intCodAlmacenOrigen = 0
    End Sub

    Private Sub dbcOrigen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen.Enter
        gStrSql = "SELECT CodAlmacenOrigen, CodAlmacenOrigen  AS DescAlmacen From CatOrigen ORDER BY DescAlmacenOrigen "
        DCGotFocus(gStrSql, dbcOrigen)
        Pon_Tool()
        mblnFueraChange = False
    End Sub

    Private Sub dbcOrigen_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dbcOrigen.KeyPress
        eventSender.keyAscii = ModEstandar.gp_CampoMayusculas(eventSender.keyAscii)
    End Sub

    Private Sub dbcOrigen_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcOrigen.KeyUp
        Dim Aux As String

        Aux = dbcOrigen.Text
        'If dbcOrigen.SelectedItem <> 0 Then
        '    dbcOrigen_Leave(dbcOrigen, New System.EventArgs())
        'End If

        dbcOrigen.Text = Aux
    End Sub

    Private Sub dbcOrigen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen.Leave
        gStrSql = "SELECT CodAlmacenOrigen, CodAlmacenOrigen AS DescAlmacen From CatOrigen  WHERE CodAlmacenOrigen LIKE '" & Trim(dbcOrigen.Text) & "%'  ORDER BY DescAlmacenOrigen "
        DCLostFocus(dbcOrigen, gStrSql, intCodAlmacenOrigen)
    End Sub

    Private Sub dbcOrigen_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcOrigen.MouseUp
        Dim Aux As String

        Aux = dbcOrigen.Text
        'If dbcOrigen.SelectedItem <> 0 Then
        '    dbcOrigen_Leave(dbcOrigen, New System.EventArgs())
        'End If

        dbcOrigen.Text = Aux
    End Sub

    '''06AGO2007 - MAVF
    Private Sub dbcSubLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSubLinea.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String


        cSubLinea = Trim(Me.dbcSubLinea.Text)

        cSubLineaDescCorta = BuscaSubLDescCortaDesc(gCODJOYERIA, mintJFam, mintJLin, Trim(Me.dbcSubLinea.Text))
        Call FormaDescripcion()

        If mblnFueraChange Then
            Exit Sub
        End If


        lStrSql = "SELECT codSubLinea, LTrim(RTrim(descSubLinea)) as descSubLinea FROM catSubLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " and codLinea = " & mintJFam & " and descSubLinea = '" & Trim(Me.dbcSubLinea.Text) & "' Order by DescSubLinea "
        ModDCombo.DCChange(lStrSql, tecla, dbcSubLinea)

        If Trim(Me.dbcSubLinea.Text) = "" Then
            mintJSub = 0
            mblnFueraChange = False
            cSubLineaDescCorta = ""
        Else

            cSubLineaDescCorta = BuscaSubLDescCortaDesc(gCODJOYERIA, mintJFam, mintJLin, Trim(Me.dbcSubLinea.Text))
        End If
        FormaDescripcion()

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcSubLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSubLinea.Enter
        Pon_Tool()
        gStrSql = "SELECT codSubLinea, LTrim(RTrim(descSubLinea)) as descSubLinea FROM catSubLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " and codLinea = " & mintJLin & " Order by DescSubLinea "
        ModDCombo.DCGotFocus(gStrSql, dbcSubLinea)
    End Sub

    Private Sub dbcSubLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSubLinea.KeyDown
        '''    If KeyCode = vbKeyEscape Then
        '''        Me.dbcLinea(0).SetFocus
        '''        KeyCode = 0
        '''    End If
        '''    tecla = KeyCode
        Select Case eventArgs.KeyCode
            Case System.Windows.Forms.Keys.Escape
                'Me._dbcLinea_0.SelectedValue(0).Focus()
                Me._dbcLinea_0.Focus()
            Case Else
                tecla = eventArgs.KeyCode
        End Select
    End Sub

    Private Sub dbcSubLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSubLinea.Leave
        '''    Dim I As Long
        '''    Dim Aux As Long
        '''    Dim cDescripcion As String
        '''
        '''    cDescripcion = Trim(Me.dbcSubLinea.text)
        '''    If Screen.ActiveForm.Name <> Me.Name Then
        '''        Exit Sub
        '''    End If
        '''    gStrSql = "SELECT codSubLinea, LTrim(RTrim(descSubLinea)) as descSubLinea FROM catSubLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " and codLinea = " & mintJLin & " and descSubLinea LIKE '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
        '''    Aux = mintJSub
        '''    mintJSub = 0
        '''    ModDCombo.DCLostFocus dbcSubLinea, gStrSql, mintJSub
        '''    cSubLinea = Trim(Me.dbcSubLinea.text)
        '''    If mintJSub = 0 Then
        '''        mblnFueraChange = True
        '''        Me.dbcSubLinea.text = cINDEFINIDA
        '''        mblnFueraChange = False
        '''        cSubLineaDescCorta = ""
        '''    Else
        '''       cSubLineaDescCorta = BuscaSubLDescCorta(gCODJOYERIA, mintJFam, mintJLin, mintJSub)
        '''    End If
        '''    FormaDescripcion


        Dim I As Integer
        Dim Aux As Integer
        Dim cDescripcion As String

        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        cDescripcion = Trim(Me.dbcSubLinea.Text)

        gStrSql = "SELECT codSubLinea, LTrim(RTrim(descSubLinea)) as descSubLinea FROM catSubLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " and codLinea = " & mintJLin & " and descSubLinea LIKE '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
        Aux = mintJSub
        mintJSub = 0
        ModDCombo.DCLostFocus(dbcSubLinea, gStrSql, mintJSub)

        cSubLinea = Trim(Me.dbcSubLinea.Text)
        If mintJSub <> Aux Then
            mblnFueraChange = True
            If mintJSub = 0 Then

                Me.dbcSubLinea.Text = cINDEFINIDA
                cSubLineaDescCorta = ""
            Else
                cSubLineaDescCorta = BuscaSubLDescCorta(gCODJOYERIA, mintJFam, mintJLin, mintJSub)
            End If
            mblnFueraChange = False
        Else
            mblnFueraChange = True
            If mintJSub = 0 Then

                Me.dbcSubLinea.Text = cINDEFINIDA
                cSubLineaDescCorta = ""
            Else
                cSubLineaDescCorta = BuscaSubLDescCorta(gCODJOYERIA, mintJFam, mintJLin, mintJSub)
            End If
            mblnFueraChange = False
        End If
        Call FormaDescripcion()

    End Sub

    Private Sub dbcSubLinea_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSubLinea.MouseUp
        Dim Aux As String

        Aux = Trim(Me.dbcSubLinea.Text)
        'If Me.dbcSubLinea.SelectedItem <> 0 Then
        'dbcSubLinea_Leave(dbcSubLinea, New System.EventArgs())
        'End If

        Me.dbcSubLinea.Text = Aux
    End Sub

    Private Sub frmCorpoABCArticulos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        BringToFront()
    End Sub

    Private Sub frmCorpoABCArticulos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmCorpoABCArticulos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Dim CtrlDwn As Boolean
        CtrlDwn = VB6.ShiftConstants.CtrlMask > 0
        If KeyCode = System.Windows.Forms.Keys.Tab Then
            If CtrlDwn Then
                KeyCode = 0
            End If
        End If
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                'If UCase(Trim(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name)) = UCase(Trim(txtCodigodelProveedor(sstArticulo.SelectedIndex).Name)) Then Exit Sub
                'If UCase(Trim(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name)) = UCase(Trim(txtDescripcion(sstArticulo.SelectedIndex).Name)) Then Exit Sub
                'If UCase(Trim(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name)) = UCase(Trim(dbcModelo.Name)) Then Exit Sub
                'If UCase(Trim(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name)) = UCase(Trim(optGenero(sstArticulo.SelectedIndex).Name)) Then Exit Sub
                'If UCase(Trim(System.Windows.Forms.Form.ActiveForm.ActiveControl.Name)) = UCase(Trim(optMoneda(sstArticulo.SelectedIndex).Name)) Then Exit Sub
                mblnFueraChange = True
                ModEstandar.AvanzarTab(Me)
                mblnFueraChange = False

            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "TXTCODARTICULO" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmCorpoABCArticulos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayúscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCorpoABCArticulos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Nuevo()
        sstArticulo.SelectedIndex = 0
        _fraContenedor_0.Enabled = True
        _fraContenedor_1.Enabled = False
        _fraContenedor_2.Enabled = False
        'En lugar de llamar al procedimiento Limpiar, Inicializa sólo las variables de importancia
        mblnNuevo = True
        mblnCambiosEnCodigo = False
        gstrNombreForma = "FRMCORPOABCARTICULOS"
    End Sub

    Private Sub frmCorpoABCArticulos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si desea cerrar la forma y esta se encuentra minimizada, esta se restaura
        If Not mblnSalir Then
            ModEstandar.RestaurarForma(Me, False)
            If Cambios() And Not (mblnNuevo) Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        If Not Guardar() Then
                            Cancel = 1
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite que se cierre el formulario
                        Cancel = 0
                    Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin Guardar
                        Cancel = 1
                End Select
            End If
        Else 'Se quiere salir con escape
            mblnSalir = False
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Sale del Formulario
                    Cancel = 0
                Case MsgBoxResult.No 'No sale del formulario
                    Me.txtCodArticulo.Focus()
                    ModEstandar.SelTextoTxt((Me.txtCodArticulo))
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCorpoABCArticulos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        gstrNombreForma = ""
    End Sub
    Private Sub optGenero_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optGenero.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = optGenero.GetIndex(eventSender)
            Select Case Index
                Case 0 'Hombre - Caballero
                    cGenero = "H"
                Case 1 'Mujer - Dama
                    cGenero = "D"
                Case 2 'Unisex - Mediano
                    cGenero = "M"
            End Select
            FormaDescripcion()
        End If
    End Sub

    Private Sub optGenero_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optGenero.Enter
        Dim Index As Integer = optGenero.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub optGenero_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles optGenero.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Dim Index As Integer = optGenero.GetIndex(eventSender)
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If Not _optMovimiento_0.Checked And Not _optMovimiento_1.Checked And Not _optMovimiento_2.Checked Then
                _optMovimiento_0.Focus()
            ElseIf _optMovimiento_0.Checked Then
                _optMovimiento_0.Focus()
            ElseIf _optMovimiento_1.Checked Then
                _optMovimiento_1.Focus()
            ElseIf _optMovimiento_2.Checked Then
                _optMovimiento_2.Focus()
            End If
        End If
    End Sub

    Private Sub optMoneda_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMoneda.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = optMoneda.GetIndex(eventSender)
            If Index > 5 Then Exit Sub
            If mblnFueraChange Then Exit Sub
            If Index = 0 Or Index = 2 Or Index = 4 Then
                cMonedaCompra = C_DOLAR
            Else
                cMonedaCompra = C_PESO
            End If
        End If
    End Sub

    Private Sub optMoneda_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles optMoneda.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Dim Index As Integer = optMoneda.GetIndex(eventSender)
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Select Case Index
                Case 10, 11
                    If Not _optMoneda_0.Checked And Not _optMoneda_1.Checked Then
                        _optMoneda_0.Focus()
                    End If
                    If _optMoneda_0.Checked Then _optMoneda_0.Focus()
                    If _optMoneda_1.Checked Then _optMoneda_1.Focus()
                Case 7, 6
                    If Not _optMoneda_2.Checked And Not _optMoneda_3.Checked Then
                        _optMoneda_2.Focus()
                    End If
                    If _optMoneda_2.Checked Then _optMoneda_2.Focus()
                    If _optMoneda_3.Checked Then _optMoneda_3.Focus()
                Case 8, 9
                    If Not _optMoneda_4.Checked And Not _optMoneda_5.Checked Then
                        _optMoneda_4.Focus()
                    End If
                    If _optMoneda_4.Checked Then _optMoneda_4.Focus()
                    If _optMoneda_5.Checked Then _optMoneda_5.Focus()
                Case 0, 1
                    txtPrecioenDolares(sstArticulo.SelectedIndex).Focus()
                Case 2, 3
                    txtPrecioenDolares(sstArticulo.SelectedIndex).Focus()
                Case 4, 5
                    txtPrecioenDolares(sstArticulo.SelectedIndex).Focus()
            End Select
        End If
    End Sub

    Private Sub optMovimiento_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMovimiento.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Integer = optMovimiento.GetIndex(eventSender)
            Select Case Index
                Case 0 'Cuarzo
                    cMovimiento = "Q"
                Case 1 'Automático
                    cMovimiento = "AUT"
                Case 2 'Manual
                    cMovimiento = "MAN"
            End Select
            FormaDescripcion()
        End If
    End Sub

    Private Sub optMovimiento_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMovimiento.Enter
        Dim Index As Integer = optMovimiento.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub sstArticulo_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sstArticulo.SelectedIndexChanged
        Static PreviousTab As Integer = sstArticulo.SelectedIndex()
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                Me.ToolTip1.SetToolTip(Me.sstArticulo, "Joyería")
                _fraContenedor_0.Enabled = True
                _fraContenedor_1.Enabled = False
                _fraContenedor_2.Enabled = False
            Case nRELOJERIA
                Me.ToolTip1.SetToolTip(Me.sstArticulo, "Relojería")
                _fraContenedor_0.Enabled = False
                _fraContenedor_1.Enabled = True
                _fraContenedor_2.Enabled = False
            Case nVARIOS
                Me.ToolTip1.SetToolTip(Me.sstArticulo, "Otros productos distintos a Joyería y Relojería")
                _fraContenedor_0.Enabled = False
                _fraContenedor_1.Enabled = False
                _fraContenedor_2.Enabled = True
        End Select
        Nuevo()
        Limpiar()
        PreviousTab = sstArticulo.SelectedIndex()
    End Sub

    Private Sub sstArticulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sstArticulo.Enter
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                Me.ToolTip1.SetToolTip(Me.sstArticulo, "Joyería")
            Case nRELOJERIA
                Me.ToolTip1.SetToolTip(Me.sstArticulo, "Relojería")
            Case nVARIOS
                Me.ToolTip1.SetToolTip(Me.sstArticulo, "Otros productos distintos a Joyería y Relojería")
        End Select
    End Sub

    Private Sub sstArticulo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles sstArticulo.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Or (KeyCode = System.Windows.Forms.Keys.Tab And Shift = 0) Then
            Select Case Me.sstArticulo.SelectedIndex
                Case nJOYERIA
                    Me.dbcFamilia.SelectedValue(0).Focus()
                Case nRELOJERIA
                    Me.dbcMarca.Focus()
                Case nVARIOS
                    Me.dbcFamilia.SelectedValue(1).Focus()
            End Select
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.txtDescArticulo.Focus()
        End If
    End Sub

    Private Sub txtAdicional_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAdicional.TextChanged
        Dim Index As Integer = txtAdicional.GetIndex(eventSender)
        FormaDescripcion()
    End Sub

    Private Sub txtAdicional_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAdicional.Enter
        Dim Index As Integer = txtAdicional.GetIndex(eventSender)
        SelTextoTxt(txtAdicional(Index))
    End Sub

    Private Sub txtAdicional_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAdicional.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Dim Index As Integer = txtAdicional.GetIndex(eventSender)
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "()[]{}.,:;#$%&=+-/\_-@")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodArtAnterior_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArtAnterior.Enter
        SelTextoTxt(txtCodArtAnterior)
    End Sub

    Private Sub txtCodArtAnterior_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodArtAnterior.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        KeyAscii = ModEstandar.MskCantidad((Me.txtCodArtAnterior.Text), KeyAscii, 5, 0, (Me.txtCodArtAnterior.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodArtAnterior_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArtAnterior.Leave
        txtCodArtAnterior.Text = CStr(CInt(Numerico((txtCodArtAnterior.Text))))
        If Trim(txtCodArtAnterior.Text) = "0" Then txtCodArtAnterior.Text = ""


        If Trim(dbcOrigen.Text) <> "" And Trim(txtCodArtAnterior.Text) <> "" And chkCodigoAnterior.CheckState Then
            If ValidaCodigoArticuloProv() Then
                If MsgBox("Este código anterior ya existe un el catálogo de artículos" & vbNewLine & vbNewLine & "Desea modificarlo???", MsgBoxStyle.Information + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton1, gstrCorpoNOMBREEMPRESA) = MsgBoxResult.Yes Then
                    dbcOrigen.Focus()
                End If
            End If
        Else
            MsgBox("Debe indicar un codigo de artículo valido", MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            dbcOrigen.Focus()
        End If
    End Sub
    Private Sub txtCodArticulo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArticulo.TextChanged
        If txtCodArticulo.Text = "" Then
            txtDescArticulo.Text = ""
            mblnNuevo = True
            Limpiar()
            mblnCambiosEnCodigo = False
        End If
        mblnCambiosEnCodigo = True
    End Sub

    Private Sub txtCodArticulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArticulo.Enter
        strControlActual = UCase("txtCodArticulo")
        SelTextoTxt((Me.txtCodArticulo))
        Pon_Tool()
    End Sub

    Private Sub txtCodArticulo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodArticulo.KeyDown
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
                    Me.txtCodArticulo.Focus()
            End Select
        End If
    End Sub

    Private Sub txtCodArticulo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodArticulo.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If (KeyAscii < System.Windows.Forms.Keys.D0 Or KeyAscii > System.Windows.Forms.Keys.D9) And KeyAscii <> System.Windows.Forms.Keys.Back Then
            KeyAscii = 0
        Else
            'Pregunta sólo si ha habido cambios
            If Cambios() And Not mblnNuevo Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes
                        If Not Guardar() Then
                            KeyAscii = 0
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite que se teclee y borre
                    Case MsgBoxResult.Cancel 'Cancela la captura
                        KeyAscii = 0
                        Me.txtCodArticulo.Focus()
                End Select
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodArticulo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodArticulo.Leave
        Dim CodAux As Integer

        'If ActiveControl.Text = Me.Text Then
        If Me.txtCodArticulo.Text <> "" Then
            ResBusquedaArt = BuscarCodigoArticulo(Trim(txtCodArticulo.Text))
            If ResBusquedaArt > 0 Or ResBusquedaArt = -1 Then
                txtCodArticulo.Text = IIf((ResBusquedaArt > 0), CInt(ResBusquedaArt), "")
                txtCodArticulo.Text = CStr(CInt(ResBusquedaArt))
                LlenaDatos()
            ElseIf ResBusquedaArt = -2 Then
                CodAux = CInt(txtCodArticulo.Text)
                txtCodArticulo.Text = ""
                BusquedaEspecial((New String("0", 6) & Trim(CStr(CodAux))))
            End If
        End If
        'End If

    End Sub

    Private Sub txtCodigodelProveedor_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigodelProveedor.TextChanged
        Dim Index As Integer = txtCodigodelProveedor.GetIndex(eventSender)
        FormaDescripcion()
    End Sub

    Private Sub txtCodigodelProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodigodelProveedor.Enter
        strControlActual = UCase("txtCodigodelProveedor")
        Dim Index As Integer = txtCodigodelProveedor.GetIndex(eventSender)
        Pon_Tool()
        SelTextoTxt(Me.txtCodigodelProveedor(Index))
    End Sub

    Private Sub txtCodigodelProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodigodelProveedor.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Dim Index As Integer = txtCodigodelProveedor.GetIndex(eventSender)
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Escape
                dbcProveedor.SelectedValue(Index).Focus()
            Case System.Windows.Forms.Keys.Return
                txtImagen(Index).Focus()
        End Select
    End Sub

    Private Sub txtCodigodelProveedor_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigodelProveedor.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Dim Index As Integer = txtCodigodelProveedor.GetIndex(eventSender)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "-/_[]()#")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    'Private Sub txtCodigodelProveedor_LostFocus(Index As Integer)
    '    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    '    If mblnFueraChange = True Then Exit Sub
    '    If Trim(txtCodigodelProveedor(Index)) = "" Then Exit Sub
    '    Dim Resultado As String
    '    Resultado = ObtenerCodArticulodeCodProv(txtCodigodelProveedor(Index))
    '    mblnFueraChange = True
    '    If Resultado > 0 Then
    '    txtCodArticulo = Resultado
    '        LlenaDatos
    '    ElseIf Resultado = -1 Then
    '        MsgBox "El código no existe." + vbNewLine + "Verifique por favor", vbExclamation + vbOKOnly, gstrCorpoNOMBREEMPRESA
    '        Exit Sub
    '    ElseIf Resultado = -2 Then
    '        BusquedaEspecial Trim(txtCodigodelProveedor(Index))
    '    End If
    '    mblnFueraChange = False
    'End Sub

    Private Sub txtCostoAdicional_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoAdicional.Enter
        Dim Index As Integer = txtCostoAdicional.GetIndex(eventSender)
        Pon_Tool()
        SelTextoTxt(Me.txtCostoAdicional(Index))
    End Sub

    Private Sub txtCostoAdicional_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCostoAdicional.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Dim Index As Integer = txtCostoAdicional.GetIndex(eventSender)
        If KeyAscii = 13 Then
            If IsNumeric(Trim(txtCostoAdicional(Index).Text)) Or Trim(txtCostoAdicional(Index).Text) = "" Then
                txtCostoAdicional(Index).Text = Format(Numerico(txtCostoAdicional(Index).Text), "###,###,##0.00")
            End If
        End If
        KeyAscii = ModEstandar.MskCantidad(txtCostoAdicional(Index).Text, KeyAscii, 9, 2, Me.txtCostoAdicional(Index).SelectionStart)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCostoAdicional_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoAdicional.Leave
        Dim Index As Integer = txtCostoAdicional.GetIndex(eventSender)
        If IsNumeric(Trim(Me.txtCostoAdicional(Index).Text)) Or Trim(Me.txtCostoAdicional(Index).Text) = "" Then
            txtCostoAdicional(Index).Text = Format(Numerico(Me.txtCostoAdicional(Index).Text), "###,###,##0.00")
            nCostoAdicional = CDec(Numerico(txtCostoAdicional(Index).Text))
            ActualizaCantidades()
        End If
    End Sub

    Private Sub txtCostoFactura_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoFactura.Enter
        Dim Index As Integer = txtCostoFactura.GetIndex(eventSender)
        Pon_Tool()
        SelTextoTxt(Me.txtCostoFactura(Index))
    End Sub

    Private Sub txtCostoFactura_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCostoFactura.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Dim Index As Integer = txtCostoFactura.GetIndex(eventSender)
        If KeyAscii = 13 Then
            If IsNumeric(Trim(Me.txtCostoFactura(Index).Text)) Or Trim(Me.txtCostoFactura(Index).Text) = "" Then
                txtCostoFactura(Index).Text = Format(Numerico(Me.txtCostoFactura(Index).Text), "###,###,##0.00")
            End If
        End If
        KeyAscii = ModEstandar.MskCantidad(Me.txtCostoFactura(Index).Text, KeyAscii, 9, 2, Me.txtCostoFactura(Index).SelectionStart)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCostoFactura_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoFactura.Leave
        Dim Index As Integer = txtCostoFactura.GetIndex(eventSender)
        If IsNumeric(Trim(Me.txtCostoFactura(Index).Text)) Or Trim(Me.txtCostoFactura(Index).Text) = "" Then
            Me.txtCostoFactura(Index).Text = Format(Numerico(Me.txtCostoFactura(Index).Text), "###,###,##0.00")
            nCostoFactura = CDec(Numerico(Me.txtCostoFactura(Index).Text))
            ActualizaCantidades()
        End If
    End Sub

    Private Sub txtCostoIndirecto_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoIndirecto.Enter
        Dim Index As Integer = txtCostoIndirecto.GetIndex(eventSender)
        Pon_Tool()
        SelTextoTxt(Me.txtCostoIndirecto(Index))
    End Sub

    Private Sub txtCostoIndirecto_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCostoIndirecto.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Dim Index As Integer = txtCostoIndirecto.GetIndex(eventSender)
        If KeyAscii = 13 Then
            If IsNumeric(Trim(Me.txtCostoIndirecto(Index).Text)) Or Trim(Me.txtCostoIndirecto(Index).Text) = "" Then
                txtCostoIndirecto(Index).Text = Format(Numerico(Me.txtCostoIndirecto(Index).Text), "###,###,##0.00")
            End If
        End If
        KeyAscii = ModEstandar.MskCantidad(Me.txtCostoIndirecto(Index).Text, KeyAscii, 9, 2, Me.txtCostoIndirecto(Index).SelectionStart)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCostoIndirecto_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoIndirecto.Leave
        Dim Index As Integer = txtCostoIndirecto.GetIndex(eventSender)
        If IsNumeric(Trim(Me.txtCostoIndirecto(Index).Text)) Or Trim(Me.txtCostoIndirecto(Index).Text) = "" Then
            Me.txtCostoIndirecto(Index).Text = Format(Numerico(Me.txtCostoIndirecto(Index).Text), "###,###,##0.00")
            nCostoIndirectos = CDec(Numerico(Me.txtCostoIndirecto(Index).Text))
            ActualizaCantidades()
        End If
    End Sub

    Private Sub txtCostoReal_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoReal.TextChanged
        Dim Index As Integer = txtCostoReal.GetIndex(eventSender)
        lblMargen(Index).Text = "0.00"
    End Sub

    Private Sub txtCostoReal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoReal.Enter
        Dim Index As Integer = txtCostoReal.GetIndex(eventSender)
        Pon_Tool()
        SelTextoTxt(Me.txtCostoReal(Index))
    End Sub

    Private Sub txtCostoReal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCostoReal.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Dim Index As Integer = txtCostoReal.GetIndex(eventSender)
        If KeyAscii = 13 Then
            txtCostoReal(Index).Text = Format(Me.txtCostoReal(Index).Text, "###,###,##0.00")
        End If
        KeyAscii = ModEstandar.MskCantidad(Me.txtCostoReal(Index).Text, KeyAscii, 9, 2, Me.txtCostoReal(Index).SelectionStart)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCostoReal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCostoReal.Leave
        Dim Index As Integer = txtCostoReal.GetIndex(eventSender)
        Dim lImporteDol As Decimal

        If Trim(txtPrecioenDolares(Index).Text) = "" Then Exit Sub
        If Not IsNumeric(DesCifrar(txtCostoReal(Index).Text)) Then Exit Sub
        If CDec(txtPrecioenDolares(Index).Text) > 0 Then
            If _optMoneda_10.Checked Then lImporteDol = (CDec(Numerico(txtPrecioenDolares(Index).Text)) / gcurCorpoTIPOCAMBIODOLAR) Else lImporteDol = CDec(Numerico(txtPrecioenDolares(Index).Text))
            lblMargen(Index).Text = Format((1 - (System.Math.Round(CDbl(DesCifrar(txtCostoReal(Index).Text)) / lImporteDol, 4))) * 100, "##0.00")
        Else
            lblMargen(Index).Text = "0.00"
        End If
        txtCostoReal(Index).Text = Format(Numerico(txtCostoReal(Index).Text), "###,###,##0.00")
    End Sub

    Private Sub txtDescArticulo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescArticulo.TextChanged
        If Trim(txtDescArticulo.Text) = "" Then txtCodArticulo.Text = ""
    End Sub

    Private Sub txtDescArticulo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescArticulo.Enter
        strControlActual = UCase("txtDescArticulo")
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub txtDescripcion_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescripcion.Enter
        Dim Index As Integer = txtDescripcion.GetIndex(eventSender)
        Pon_Tool()
        ModEstandar.SelTextoTxt(Me.txtDescripcion(Index))
    End Sub

    Private Sub txtDescripcion_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDescripcion.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Dim Index As Integer = txtDescripcion.GetIndex(eventSender)
        Dim vlIndex As Integer

        If KeyCode = System.Windows.Forms.Keys.Return Then
            Select Case Index
                Case 0 : vlIndex = 10
                Case 1 : vlIndex = 7
                Case 2 : vlIndex = 8
            End Select
            optMoneda(vlIndex).Focus()
        End If
        If KeyCode = System.Windows.Forms.Keys.Escape Then txtAdicional(Index).Focus()
    End Sub

    Private Sub txtMDSCertificado_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSCertificado.Enter
        SelTextoTxt(txtMDSCertificado)
    End Sub

    Private Sub txtMDSCertificado_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMDSCertificado.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "-")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtMDSColor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSColor.Enter
        SelTextoTxt(txtMDSColor)
    End Sub

    Private Sub txtMDSColor_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMDSColor.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoLetras(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtMDSPeso_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSPeso.TextChanged
        If CDec(ModEstandar.Numerico((txtMDSPeso.Text))) > 100 Then
            MsgBox("Valor incorrecto" & vbNewLine & "El peso no debe pasar de 100.00" & vbNewLine & vbNewLine & "Vefifique por favor...", MsgBoxStyle.Exclamation, gstrCorpoNOMBREEMPRESA)
            txtMDSPeso.Focus()
        End If
    End Sub

    Private Sub txtMDSPeso_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSPeso.Enter
        SelTextoTxt(txtMDSPeso)
    End Sub

    Private Sub txtMDSPeso_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtMDSPeso.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = 13 Then txtMDSPeso.Text = Format(ModEstandar.Numerico((txtMDSPeso.Text)), "##0.00")
    End Sub

    Private Sub txtMDSPeso_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMDSPeso.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii, ".")
        KeyAscii = ModEstandar.MskCantidad((txtMDSPeso.Text), KeyAscii, 3, 2, (txtMDSPeso.SelectionStart))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtMDSPureza_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMDSPureza.Enter
        SelTextoTxt(txtMDSPureza)
    End Sub

    Private Sub txtMDSPureza_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMDSPureza.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPrecioenDolares_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPrecioenDolares.TextChanged
        Dim Index As Integer = txtPrecioenDolares.GetIndex(eventSender)
        lblMargen(Index).Text = "0.00"
    End Sub

    Private Sub txtPrecioenDolares_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPrecioenDolares.Enter
        Dim Index As Integer = txtPrecioenDolares.GetIndex(eventSender)
        Pon_Tool()
        SelTextoTxt(Me.txtPrecioenDolares(Index))
    End Sub

    Private Sub txtPrecioenDolares_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPrecioenDolares.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        Dim Index As Integer = txtPrecioenDolares.GetIndex(eventSender)
        If KeyAscii = 13 Then
            Me.txtPrecioenDolares(Index).Text = Format(Numerico(Me.txtPrecioenDolares(Index).Text), "###,###,##0.00")
        End If
        KeyAscii = ModEstandar.MskCantidad(Me.txtPrecioenDolares(Index).Text, KeyAscii, 9, 2, Me.txtPrecioenDolares(Index).SelectionStart)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPrecioenDolares_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPrecioenDolares.Leave
        Dim Index As Integer = txtPrecioenDolares.GetIndex(eventSender)
        Dim lImporteD As Decimal

        If Trim(txtPrecioenDolares(Index).Text) = "" Then Exit Sub
        If Not IsNumeric(DesCifrar(txtCostoReal(Index).Text)) Then Exit Sub

        If CDec(Numerico(txtPrecioenDolares(Index).Text)) > 0 Then
            If sstArticulo.SelectedIndex = 0 Then
                If _optMoneda_10.Checked Then lImporteD = CDec(Numerico(txtPrecioenDolares(Index).Text)) Else lImporteD = (CDec(Numerico(txtPrecioenDolares(Index).Text)) / gcurCorpoTIPOCAMBIODOLAR)
                lblMargen(Index).Text = Format((1 - (System.Math.Round(CDec(DesCifrar(txtCostoReal(Index).Text)) / lImporteD, 4))) * 100, "##0.00")
            ElseIf sstArticulo.SelectedIndex = 1 Then
                If _optMoneda_7.Checked Then lImporteD = CDec(Numerico(txtPrecioenDolares(Index).Text)) Else lImporteD = (CDec(Numerico(txtPrecioenDolares(Index).Text)) / gcurCorpoTIPOCAMBIODOLAR)
                lblMargen(Index).Text = Format((1 - (System.Math.Round(CDec(DesCifrar(txtCostoReal(Index).Text)) / lImporteD, 4))) * 100, "##0.00")
            ElseIf sstArticulo.SelectedIndex = 2 Then
                If _optMoneda_8.Checked Then lImporteD = CDec(Numerico(txtPrecioenDolares(Index).Text)) Else lImporteD = (CDec(Numerico(txtPrecioenDolares(Index).Text)) / gcurCorpoTIPOCAMBIODOLAR)
                lblMargen(Index).Text = Format((1 - (System.Math.Round(CDec(DesCifrar(txtCostoReal(Index).Text)) / lImporteD, 4))) * 100, "##0.00")
            End If
        Else
            lblMargen(Index).Text = "0.00"
        End If

        txtPrecioenDolares(Index).Text = Format(Numerico(txtPrecioenDolares(Index).Text), "###,###,##0.00")
    End Sub

    Private Sub DefineCondicionesJoy(ByRef nFam As String, ByRef nLin As String, ByRef nSubL As String, ByRef nKil As String, ByRef nTipoM As String)
        If mintJFam <> 0 Then nFam = " CodFamilia = " & mintJFam Else nFam = " CodFamilia Is Null "
        If mintJLin <> 0 Then nLin = " CodLinea = " & mintJLin Else nLin = " CodLinea Is Null "
        If mintJSub <> 0 Then nSubL = " CodSubLinea = " & mintJSub Else nSubL = " CodSubLinea Is Null "
        If mintCodKilates <> 0 Then nKil = " CodKilates = " & mintCodKilates Else nKil = " CodKilates Is Null "
        If mintJMaterial <> 0 Then nTipoM = " CodTipoMaterial = " & mintJMaterial Else nTipoM = " CodTipoMaterial Is Null "
    End Sub

    Private Sub DefineCondicionesRel(ByRef nMar As String, ByRef nMod As String, ByRef nGen As String, ByRef nMov As String, ByRef nCrono As String, ByRef nTipoM As String)
        If mintRMar <> 0 Then nMar = " CodMarca = " & mintRMar Else nMar = " CodMarca Is Null "
        If mintRMod <> 0 Then nMod = " CodModelo = " & mintRMod Else nMod = " CodModelo Is Null "
        If mintRMaterial <> 0 Then nTipoM = " CodTipoMaterial = " & mintRMaterial Else nTipoM = " CodTipoMaterial Is Null "
        nGen = " Genero = '" & cGenero & "'"
        nMov = " Movimiento = '" & cMovimiento & "'"
        If Not chkCrono.CheckState Then nCrono = " Crono = 0 " Else nCrono = " Crono = 1 "
    End Sub

    Private Sub DefineCondicionesVar(ByRef nFam As String, ByRef nLin As String, ByRef nTipoM As String)
        If mintVFam <> 0 Then nFam = " CodFamilia = " & mintVFam Else nFam = " CodFamilia Is Null "
        If mintVLin <> 0 Then nLin = " CodLinea = " & mintVLin Else nLin = " CodLinea Is Null "
        If mintVMaterial <> 0 Then nTipoM = " CodTipoMaterial = " & mintVMaterial Else nTipoM = " CodTipoMaterial Is Null "
    End Sub

    Private Function ValidaCodigoArticuloProv() As Boolean
        Dim Rs As ADODB.Recordset
        gStrSql = "Select * From CatArticulos (Nolock) Where OrigenAnt = " & Trim(dbcOrigen.Text) & " And CodigoAnt = " & Trim(txtCodArtAnterior.Text)
        If Not mblnNuevo Then
            gStrSql = gStrSql & " And CodArticulo <> " & Trim(txtCodArticulo.Text)
        End If

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        Rs = Cmd.Execute

        If Rs.RecordCount > 0 Then
            ValidaCodigoArticuloProv = True
        Else
            ValidaCodigoArticuloProv = False
        End If

    End Function

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnGuardar_Click_1(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Guardar()
    End Sub

    Private Sub btnEliminar_Click_1(sender As Object, e As EventArgs) Handles btnEliminar.Click
        Eliminar()
    End Sub

    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        Nuevo()
    End Sub


    '    Private Sub dbcFamilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcFamilia.Enter
    '        Dim Index As Integer = dbcFamilia.SelectedValue(eventSender)
    '        Pon_Tool()
    '        Select Case Index
    '            Case 0 'JOYERIA
    '                gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " ORDER BY descFamilia "
    '            Case 1 'VARIOS
    '                gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODVARIOS & " ORDER BY descFamilia "
    '        End Select
    '        ModDCombo.DCGotFocus(gStrSql, dbcFamilia.SelectedValue(Index))
    '    End Sub

    '    Private Sub dbcFamilia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcFamilia.KeyDown
    '        Dim Index As Integer = dbcFamilia.SelectedValue(eventSender)
    '        Select Case eventArgs.KeyCode
    '            Case System.Windows.Forms.Keys.Escape
    '                Me.sstArticulo.Focus()
    '            Case Else
    '                Select Case Index
    '                    Case 0 'JOYERIA
    '                        tecla = eventArgs.KeyCode
    '                    Case 1 'VARIOS
    '                        tecla = eventArgs.KeyCode
    '                End Select
    '        End Select
    '    End Sub

    '    Private Sub dbcFamilia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcFamilia.Leave
    '        Dim Index As Integer = dbcFamilia.SelectedValue(eventSender)
    '        Dim I As Integer
    '        Dim Aux As Integer 'Almacena el anterior
    '        Dim cDescripcion As String
    '        Dim cLIKE As String

    '        cDescripcion = Trim(Me.dbcFamilia.SelectedValue(Index).Text)
    '        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
    '        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
    '        '    Exit Sub
    '        'End If
    '        Select Case Index
    '            Case 0 'JOYERIA
    '                gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
    '                Aux = mintJFam
    '                mintJFam = 0
    '                ModDCombo.DCLostFocus(dbcFamilia.SelectedValue(Index), gStrSql, mintJFam)

    '                cFamilia(Index) = Trim(Me.dbcFamilia.SelectedValue(Index).Text)
    '                If mintJFam <> Aux Then
    '                    mblnFueraChange = True
    '                    If mintJFam = 0 Then

    '                        Me.dbcFamilia.SelectedValue(Index).Text = cINDEFINIDA
    '                    End If
    '                    mintJLin = 0

    '                    Me.dbcLinea.SelectedValue(Index).Text = cINDEFINIDA
    '                    mintJSub = 0

    '                    Me.dbcSubLinea.Text = cINDEFINIDA
    '                    mblnFueraChange = False
    '                Else 'Si no cambió, y es indefinido
    '                    If mintJFam = 0 Then
    '                        mblnFueraChange = True

    '                        Me.dbcFamilia.SelectedValue(Index).Text = cINDEFINIDA
    '                        mblnFueraChange = False
    '                    End If
    '                End If
    '                Call FormaDescripcion()
    '            Case 1 'VARIOS
    '                gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODVARIOS & " and descFamilia = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
    '                Aux = mintVFam
    '                mintVFam = 0
    '                ModDCombo.DCLostFocus(dbcFamilia.SelectedValue(Index), gStrSql, mintVFam)

    '                cFamilia(Index) = Trim(Me.dbcFamilia.SelectedValue(Index).Text)
    '                If mintVFam <> Aux Then
    '                    mblnFueraChange = True
    '                    If mintVFam = 0 Then

    '                        Me.dbcFamilia.SelectedValue(Index).Text = cINDEFINIDA
    '                    End If
    '                    mintVLin = 0

    '                    Me.dbcLinea.SelectedValue(Index).Text = cINDEFINIDA
    '                    mblnFueraChange = False
    '                Else
    '                    If mintVFam = 0 Then
    '                        mblnFueraChange = True

    '                        Me.dbcFamilia.SelectedValue(Index).Text = cINDEFINIDA
    '                        mblnFueraChange = False
    '                    End If
    '                End If
    '                Call FormaDescripcion()
    '        End Select
    '    End Sub

    '    Private Sub dbcFamilia_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcFamilia.MouseUp
    '        Dim Index As Integer = dbcFamilia.SelectedValue(eventSender)
    '        Dim Aux As String

    '        Aux = Trim(Me.dbcFamilia.SelectedValue(Index).Text)
    '        'If Me.dbcFamilia.SelectedValue(Index).SelectedItem <> 0 Then
    '        'dbcFamilia_Leave(dbcFamilia.SelectedValue.Item(Index), New System.EventArgs())
    '        'End If

    '        Me.dbcFamilia.SelectedValue(Index).Text = Aux
    '    End Sub

    '    Private Sub dbcFamilia_CursorChanged(sender As Object, e As EventArgs) Handles dbcFamilia.CursorChanged
    '        Dim Index As Integer = dbcFamilia.SelectedValue(sender)
    '        On Error GoTo MError
    '        Dim lStrSql As String


    '        cFamilia(Index) = Trim(Me.dbcFamilia.SelectedValue(Index).Text)
    '        Call FormaDescripcion()

    '        If mblnFueraChange Then
    '            Exit Sub
    '        End If

    '        Select Case Index
    '            Case 0 'JOYERIA

    '                lStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia LIKE '" & Trim(Me.dbcFamilia.SelectedValue(Index).Text) & "%' Order by DescFamilia "
    '                ModDCombo.DCChange(lStrSql, tecla, dbcFamilia.SelectedValue(Index))
    '            Case 1 'VARIOS

    '                lStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODVARIOS & " and descFamilia LIKE '" & Trim(Me.dbcFamilia.SelectedValue(Index).Text) & "%' Order by DescFamilia "
    '                ModDCombo.DCChange(lStrSql, tecla, dbcFamilia.SelectedValue(Index))
    '        End Select


    '        If Trim(Me.dbcFamilia.SelectedValue(Index).Text) = "" Then
    '            mblnFueraChange = True
    '            mintJFam = 0
    '            mintVFam = 0
    '            mintJLin = 0
    '            mintVLin = 0

    '            Me._dbcLinea_0.SelectedValue(Index).Text = cINDEFINIDA
    '            If Index = 0 Then
    '                mintJSub = 0

    '                Me.dbcSubLinea.Text = cINDEFINIDA
    '            End If
    '            mblnFueraChange = True
    '        End If

    'MError:
    '        If Err.Number <> 0 Then
    '            ModEstandar.MostrarError()
    '        End If
    '    End Sub



    Private Sub _dbcFamilia_0_CursorChanged(sender As Object, e As EventArgs) Handles _dbcFamilia_0.CursorChanged
        Dim Index As Integer
        '= _dbcFamilia_0.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String


        cFamilia(Index) = Trim(Me._dbcFamilia_0.SelectedValue(Index).Text)
        Call FormaDescripcion()

        If mblnFueraChange Then
            Exit Sub
        End If

        Select Case Index
            Case 0 'JOYERIA

                lStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia LIKE '" & Trim(Me._dbcFamilia_0.SelectedValue(Index).Text) & "%' Order by DescFamilia "
                ModDCombo.DCChange(lStrSql, tecla, _dbcFamilia_0.SelectedValue(Index))
            Case 1 'VARIOS

                lStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODVARIOS & " and descFamilia LIKE '" & Trim(Me._dbcFamilia_0.SelectedValue(Index).Text) & "%' Order by DescFamilia "
                ModDCombo.DCChange(lStrSql, tecla, _dbcFamilia_0.SelectedValue(Index))
        End Select


        If Trim(Me._dbcFamilia_0.SelectedValue(Index).Text) = "" Then
            mblnFueraChange = True
            mintJFam = 0
            mintVFam = 0
            mintJLin = 0
            mintVLin = 0

            Me._dbcLinea_0.SelectedValue(Index).Text = cINDEFINIDA
            If Index = 0 Then
                mintJSub = 0

                Me.dbcSubLinea.Text = cINDEFINIDA
            End If
            mblnFueraChange = True
        End If

MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub _dbcFamilia_0_Enter(sender As Object, e As EventArgs) Handles _dbcFamilia_0.Enter
        Dim Index As Integer
        '= _dbcFamilia_0.SelectedItem(sender)
        Pon_Tool()
        Select Case Index
            Case 0 'JOYERIA
                gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " ORDER BY descFamilia "
            Case 1 'VARIOS
                gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODVARIOS & " ORDER BY descFamilia "
        End Select
        'ModDCombo.DCGotFocus(gStrSql, _dbcFamilia_0.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _dbcFamilia_0)
    End Sub

    Private Sub _dbcFamilia_0_Leave(sender As Object, e As EventArgs) Handles _dbcFamilia_0.Leave
        Dim Index As Integer
        '= _dbcFamilia_0.SelectedValue(sender)
        Dim I As Integer
        Dim Aux As Integer 'Almacena el anterior
        Dim cDescripcion As String
        Dim cLIKE As String

        'cDescripcion = Trim(Me._dbcFamilia_0.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._dbcFamilia_0.Text)
        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        Select Case Index
            Case 0 'JOYERIA
                gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
                Aux = mintJFam
                mintJFam = 0
                'ModDCombo.DCLostFocus(_dbcFamilia_0.SelectedValue(Index), gStrSql, mintJFam)
                ModDCombo.DCLostFocus(_dbcFamilia_0, gStrSql, mintJFam)

                'cFamilia(Index) = Trim(Me._dbcFamilia_0.SelectedValue(Index).Text)
                cFamilia(Index) = Trim(Me._dbcFamilia_0.Text)
                If mintJFam <> Aux Then
                    mblnFueraChange = True
                    If mintJFam = 0 Then

                        'Me._dbcFamilia_0.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcFamilia_0.Text = cINDEFINIDA
                    End If
                    mintJLin = 0

                    'Me._dbcFamilia_0.SelectedValue(Index).Text = cINDEFINIDA
                    Me._dbcFamilia_0.Text = cINDEFINIDA
                    mintJSub = 0

                    Me._dbcFamilia_0.Text = cINDEFINIDA
                    mblnFueraChange = False
                Else 'Si no cambió, y es indefinido
                    If mintJFam = 0 Then
                        mblnFueraChange = True

                        'Me._dbcFamilia_0.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcFamilia_0.Text = cINDEFINIDA
                        mblnFueraChange = False
                    End If
                End If
                Call FormaDescripcion()
            Case 1 'VARIOS
                gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODVARIOS & " and descFamilia = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
                Aux = mintVFam
                mintVFam = 0
                'ModDCombo.DCLostFocus(_dbcFamilia_0.SelectedValue(Index), gStrSql, mintVFam)
                ModDCombo.DCLostFocus(_dbcFamilia_0, gStrSql, mintVFam)

                'cFamilia(Index) = Trim(Me._dbcFamilia_0.SelectedValue(Index).Text)
                cFamilia(Index) = Trim(Me._dbcFamilia_0.Text)
                If mintVFam <> Aux Then
                    mblnFueraChange = True
                    If mintVFam = 0 Then

                        'Me._dbcFamilia_0.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcFamilia_0.Text = cINDEFINIDA
                    End If
                    mintVLin = 0

                    'Me._dbcLinea_0.SelectedValue(Index).Text = cINDEFINIDA
                    Me._dbcFamilia_0.Text = cINDEFINIDA
                    mblnFueraChange = False
                Else
                    If mintVFam = 0 Then
                        mblnFueraChange = True

                        'Me._dbcFamilia_0.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcFamilia_0.Text = cINDEFINIDA
                        mblnFueraChange = False
                    End If
                End If
                Call FormaDescripcion()
        End Select
    End Sub

    Private Sub _dbcFamilia_0_MouseUp(sender As Object, e As MouseEventArgs) Handles _dbcFamilia_0.MouseUp
        Dim Index As Integer
        '= _dbcFamilia_0.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcFamilia_0.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcFamilia_0.Text)
        'If Me._dbcFamilia_0.SelectedValue(Index).SelectedItem <> 0 Then
        '_dbcFamilia_0_Leave(_dbcFamilia_0.SelectedValue.Item(Index), New System.EventArgs())
        'End If

        'Me._dbcFamilia_0.SelectedValue(Index).Text = Aux
        Me._dbcFamilia_0.Text = Aux
    End Sub

    Private Sub _dbcFamilia_0_KeyDown(sender As Object, e As KeyEventArgs) Handles _dbcFamilia_0.KeyDown
        Dim Index As Integer = _dbcFamilia_0.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Me.sstArticulo.Focus()
            Case Else
                Select Case Index
                    Case 0 'JOYERIA
                        tecla = e.KeyCode
                    Case 1 'VARIOS
                        tecla = e.KeyCode
                End Select
        End Select
    End Sub



    Private Sub _dbcFamilia_1_CursorChanged(sender As Object, e As EventArgs) Handles _dbcFamilia_1.CursorChanged
        Dim Index As Integer
        '= _dbcFamilia_1.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String


        'cFamilia(Index) = Trim(Me._dbcFamilia_1.SelectedValue(Index).Text)
        cFamilia(Index) = Trim(Me._dbcFamilia_1.Text)

        Call FormaDescripcion()

        If mblnFueraChange Then
            Exit Sub
        End If

        Select Case Index
            Case 0 'JOYERIA

                lStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia LIKE '" & Trim(Me._dbcFamilia_1.Text) & "%' Order by DescFamilia "
                ModDCombo.DCChange(lStrSql, tecla, _dbcFamilia_1)
            Case 1 'VARIOS

                lStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODVARIOS & " and descFamilia LIKE '" & Trim(Me._dbcFamilia_1.Text) & "%' Order by DescFamilia "
                ModDCombo.DCChange(lStrSql, tecla, _dbcFamilia_1)
        End Select


        If Trim(Me._dbcFamilia_1.Text) = "" Then
            mblnFueraChange = True
            mintJFam = 0
            mintVFam = 0
            mintJLin = 0
            mintVLin = 0

            Me._dbcLinea_1.Text = cINDEFINIDA
            If Index = 0 Then
                mintJSub = 0

                Me.dbcSubLinea.Text = cINDEFINIDA
            End If
            mblnFueraChange = True
        End If

MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub _dbcFamilia_1_Enter(sender As Object, e As EventArgs) Handles _dbcFamilia_1.Enter
        Dim Index As Integer
        '= _dbcFamilia_1.SelectedItem(sender)
        Pon_Tool()
        Select Case Index
            Case 0 'JOYERIA
                gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " ORDER BY descFamilia "
            Case 1 'VARIOS
                gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODVARIOS & " ORDER BY descFamilia "
        End Select
        'ModDCombo.DCGotFocus(gStrSql, _dbcFamilia_1.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _dbcFamilia_1)
    End Sub

    Private Sub _dbcFamilia_1_Leave(sender As Object, e As EventArgs) Handles _dbcFamilia_1.Leave
        Dim Index As Integer
        '= _dbcFamilia_1.SelectedValue(sender)
        Dim I As Integer
        Dim Aux As Integer 'Almacena el anterior
        Dim cDescripcion As String
        Dim cLIKE As String

        'cDescripcion = Trim(Me._dbcFamilia_1.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._dbcFamilia_1.Text)
        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        Select Case Index
            Case 0 'JOYERIA
                gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODJOYERIA & " and descFamilia = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
                Aux = mintJFam
                mintJFam = 0
                'ModDCombo.DCLostFocus(_dbcFamilia_1.SelectedValue(Index), gStrSql, mintJFam)
                ModDCombo.DCLostFocus(_dbcFamilia_1, gStrSql, mintJFam)

                'cFamilia(Index) = Trim(Me._dbcFamilia_1.SelectedValue(Index).Text)
                cFamilia(Index) = Trim(Me._dbcFamilia_1.Text)
                If mintJFam <> Aux Then
                    mblnFueraChange = True
                    If mintJFam = 0 Then

                        'Me._dbcFamilia_1.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcFamilia_1.Text = cINDEFINIDA
                    End If
                    mintJLin = 0

                    'Me._dbcFamilia_1.SelectedValue(Index).Text = cINDEFINIDA
                    Me._dbcFamilia_1.Text = cINDEFINIDA
                    mintJSub = 0

                    Me._dbcFamilia_1.Text = cINDEFINIDA
                    mblnFueraChange = False
                Else 'Si no cambió, y es indefinido
                    If mintJFam = 0 Then
                        mblnFueraChange = True

                        'Me._dbcFamilia_1.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcFamilia_1.Text = cINDEFINIDA
                        mblnFueraChange = False
                    End If
                End If
                Call FormaDescripcion()
            Case 1 'VARIOS
                gStrSql = "SELECT codFamilia, RTrim(LTrim(descFamilia)) as descFamilia FROM catFamilias Where codGrupo = " & gCODVARIOS & " and descFamilia = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
                Aux = mintVFam
                mintVFam = 0
                'ModDCombo.DCLostFocus(_dbcFamilia_1.SelectedValue(Index), gStrSql, mintVFam)
                ModDCombo.DCLostFocus(_dbcFamilia_1, gStrSql, mintVFam)

                'cFamilia(Index) = Trim(Me._dbcFamilia_1.SelectedValue(Index).Text)
                cFamilia(Index) = Trim(Me._dbcFamilia_1.Text)
                If mintVFam <> Aux Then
                    mblnFueraChange = True
                    If mintVFam = 0 Then

                        'Me._dbcFamilia_1.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcFamilia_1.Text = cINDEFINIDA
                    End If
                    mintVLin = 0

                    'Me._dbcFamilia_1.SelectedValue(Index).Text = cINDEFINIDA
                    Me._dbcFamilia_1.Text = cINDEFINIDA
                    mblnFueraChange = False
                Else
                    If mintVFam = 0 Then
                        mblnFueraChange = True

                        'Me._dbcFamilia_1.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcFamilia_1.Text = cINDEFINIDA
                        mblnFueraChange = False
                    End If
                End If
                Call FormaDescripcion()
        End Select
    End Sub

    Private Sub _dbcFamilia_1_MouseUp(sender As Object, e As MouseEventArgs) Handles _dbcFamilia_1.MouseUp
        Dim Index As Integer
        '= _dbcFamilia_1.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcFamilia_1.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcFamilia_1.Text)
        'If Me._dbcFamilia_1.SelectedValue(Index).SelectedItem <> 0 Then
        '_dbcFamilia_1_Leave(_dbcFamilia_1.SelectedValue.Item(Index), New System.EventArgs())
        'End If

        'Me._dbcFamilia_1.SelectedValue(Index).Text = Aux
        Me._dbcFamilia_1.Text = Aux
    End Sub

    Private Sub _dbcFamilia_1_KeyDown(sender As Object, e As KeyEventArgs) Handles _dbcFamilia_1.KeyDown
        Dim Index As Integer
        '= _dbcFamilia_1.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Me.sstArticulo.Focus()
            Case Else
                Select Case Index
                    Case 0 'JOYERIA
                        tecla = e.KeyCode
                    Case 1 'VARIOS
                        tecla = e.KeyCode
                End Select
        End Select
    End Sub



    '    Private Sub dbcLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcLinea.CursorChanged
    '        Dim Index As Integer = dbcLinea.SelectedValue(eventSender)
    '        On Error GoTo MError
    '        Dim lStrSql As String


    '        cLinea(Index) = Trim(Me.dbcLinea.SelectedValue(Index).Text)
    '        Call FormaDescripcion()

    '        If mblnFueraChange Then
    '            Exit Sub
    '        End If

    '        Select Case Index
    '            Case 0 'JOYERIA

    '                lStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " and descLinea = '" & Trim(Me.dbcLinea.SelectedValue(Index).Text) & "' Order by DescLinea "
    '                ModDCombo.DCChange(lStrSql, tecla, dbcLinea.SelectedValue(Index))
    '            Case 1 'VARIOS

    '                lStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODVARIOS & " and codFamilia = " & mintVFam & " and descLinea = '" & Trim(Me.dbcLinea.SelectedValue(Index).Text) & "' Order by DescLinea "
    '                ModDCombo.DCChange(lStrSql, tecla, dbcLinea.SelectedValue(Index))
    '        End Select


    '        If Trim(Me.dbcLinea.SelectedValue(Index).Text) = "" Then
    '            mintJLin = 0
    '            mintVLin = 0
    '            If Index = 0 Then
    '                mblnFueraChange = True
    '                mintJSub = 0

    '                Me.dbcSubLinea.Text = cINDEFINIDA
    '                mblnFueraChange = False
    '            End If
    '        End If
    '        Exit Sub
    'MError:
    '        ModEstandar.MostrarError()
    '    End Sub

    '    Private Sub dbcLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcLinea.Enter
    '        Dim Index As Integer = dbcLinea.SelectedValue(eventSender)
    '        Pon_Tool()
    '        Select Case Index
    '            Case 0 'JOYERIA
    '                gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " ORDER BY descLinea "
    '            Case 1 'VARIOS
    '                gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODVARIOS & " and codFamilia = " & mintVFam & " ORDER BY descLinea "
    '        End Select
    '        ModDCombo.DCGotFocus(gStrSql, dbcLinea.SelectedValue(Index))
    '    End Sub

    '    Private Sub dbcLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcLinea.KeyDown
    '        Dim Index As Integer = dbcLinea.SelectedValue(eventSender)
    '        Select Case eventArgs.KeyCode
    '            Case System.Windows.Forms.Keys.Escape
    '                Me.dbcFamilia.SelectedValue(Index).Focus()
    '            Case Else
    '                tecla = eventArgs.KeyCode
    '        End Select
    '    End Sub

    '    Private Sub dbcLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcLinea.Leave
    '        Dim Index As Integer = dbcLinea.SelectedValue(eventSender)
    '        Dim I As Integer
    '        Dim Aux As Integer
    '        Dim cDescripcion As String
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
    '            Exit Sub
    '        End If

    '        cDescripcion = Trim(Me.dbcLinea.SelectedValue(Index).Text)
    '        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
    '        Select Case Index
    '            Case 0 'JOYERIA
    '                gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " and descLinea = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
    '                Aux = mintJLin
    '                mintJLin = 0
    '                ModDCombo.DCLostFocus(dbcLinea.SelectedValue(Index), gStrSql, mintJLin)

    '                cLinea(Index) = Trim(Me.dbcLinea.SelectedValue(Index).Text)
    '                If mintJLin <> Aux Then
    '                    mblnFueraChange = True
    '                    If mintJLin = 0 Then

    '                        Me.dbcLinea.SelectedValue(Index).Text = cINDEFINIDA
    '                    End If
    '                    mintJSub = 0

    '                    Me.dbcSubLinea.Text = cINDEFINIDA
    '                    mblnFueraChange = False
    '                Else
    '                    mblnFueraChange = True
    '                    If mintJLin = 0 Then

    '                        Me.dbcLinea.SelectedValue(Index).Text = cINDEFINIDA
    '                    End If
    '                    mblnFueraChange = False
    '                End If
    '                Call FormaDescripcion()
    '            Case 1 'VARIOS
    '                gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODVARIOS & " and codFamilia = " & mintVFam & " and descLinea = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
    '                Aux = mintVLin
    '                mintVLin = 0
    '                ModDCombo.DCLostFocus(dbcLinea.SelectedValue(Index), gStrSql, mintVLin)

    '                cLinea(Index) = Trim(Me.dbcLinea.SelectedValue(Index).Text)
    '                If Aux <> mintVLin Then
    '                    If mintVLin = 0 Then
    '                        mblnFueraChange = True

    '                        Me.dbcLinea.SelectedValue(Index).Text = cINDEFINIDA
    '                        mblnFueraChange = False
    '                    End If
    '                Else
    '                    mblnFueraChange = True
    '                    If mintVLin = 0 Then

    '                        Me.dbcLinea.SelectedValue(Index).Text = cINDEFINIDA
    '                    End If
    '                    mblnFueraChange = False
    '                End If
    '                Call FormaDescripcion()
    '        End Select
    '    End Sub

    '    Private Sub dbcLinea_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcLinea.MouseUp
    '        Dim Index As Integer = dbcLinea.SelectedValue(eventSender)
    '        Dim Aux As String

    '        Aux = Trim(Me.dbcLinea.SelectedValue(Index).Text)
    '        If Me.dbcLinea.SelectedValue(Index).SelectedItem <> 0 Then
    '            dbcLinea_Leave(dbcLinea.SelectedValue.Item(Index), New System.EventArgs())
    '        End If

    '        Me.dbcLinea.SelectedValue(Index).Text = Aux
    '    End Sub


    Private Sub _dbcLinea_0_KeyDown(sender As Object, e As KeyEventArgs) Handles _dbcLinea_0.KeyDown
        Dim Index As Integer
        '= _dbcLinea_0.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                'Me._dbcFamilia_0.SelectedValue(Index).Focus()
                Me._dbcFamilia_0.Focus()
            Case Else
                tecla = e.KeyCode
        End Select
    End Sub

    Private Sub _dbcLinea_0_MouseUp(sender As Object, e As MouseEventArgs) Handles _dbcLinea_0.MouseUp
        Dim Index As Integer
        '= _dbcLinea_0.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcLinea_0.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcLinea_0.Text)
        'If Me._dbcLinea_0.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._dbcLinea_0.SelectedItem <> 0 Then
        '_dbcLinea_0_Leave(_dbcLinea_0.SelectedValue.Item(Index), New System.EventArgs())
        '_dbcLinea_0_Leave(_dbcLinea_0, New System.EventArgs())
        'End If

        'Me._dbcLinea_0.SelectedValue(Index).Text = Aux
        Me._dbcLinea_0.Text = Aux
    End Sub

    Private Sub _dbcLinea_0_Enter(sender As Object, e As EventArgs) Handles _dbcLinea_0.Enter
        Dim Index As Integer
        '= _dbcLinea_0.SelectedValue(sender)
        Pon_Tool()
        Select Case Index
            Case 0 'JOYERIA
                gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " ORDER BY descLinea "
            Case 1 'VARIOS
                gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODVARIOS & " and codFamilia = " & mintVFam & " ORDER BY descLinea "
        End Select
        'ModDCombo.DCGotFocus(gStrSql, _dbcLinea_0.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _dbcLinea_0)
    End Sub

    Private Sub _dbcLinea_0_Leave(sender As Object, e As EventArgs) Handles _dbcLinea_0.Leave
        Dim Index As Integer
        '= _dbcLinea_0.SelectedValue(sender)
        Dim I As Integer
        Dim Aux As Integer
        Dim cDescripcion As String
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        'cDescripcion = Trim(Me._dbcLinea_0.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._dbcLinea_0.Text)
        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
        Select Case Index
            Case 0 'JOYERIA
                gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " and descLinea = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
                Aux = mintJLin
                mintJLin = 0
                'ModDCombo.DCLostFocus(_dbcLinea_0.SelectedValue(Index), gStrSql, mintJLin)
                ModDCombo.DCLostFocus(_dbcLinea_0, gStrSql, mintJLin)

                'cLinea(Index) = Trim(Me._dbcLinea_0.SelectedValue(Index).Text)
                cLinea(Index) = Trim(Me._dbcLinea_0.Text)
                If mintJLin <> Aux Then
                    mblnFueraChange = True
                    If mintJLin = 0 Then

                        'Me._dbcLinea_0.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcLinea_0.Text = cINDEFINIDA
                    End If
                    mintJSub = 0

                    Me.dbcSubLinea.Text = cINDEFINIDA
                    mblnFueraChange = False
                Else
                    mblnFueraChange = True
                    If mintJLin = 0 Then

                        'Me._dbcLinea_0.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcLinea_0.Text = cINDEFINIDA
                    End If
                    mblnFueraChange = False
                End If
                Call FormaDescripcion()
            Case 1 'VARIOS
                gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODVARIOS & " and codFamilia = " & mintVFam & " and descLinea = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
                Aux = mintVLin
                mintVLin = 0
                'ModDCombo.DCLostFocus(_dbcLinea_0.SelectedValue(Index), gStrSql, mintVLin)
                ModDCombo.DCLostFocus(_dbcLinea_0, gStrSql, mintVLin)

                'cLinea(Index) = Trim(Me._dbcLinea_0.SelectedValue(Index).Text)
                cLinea(Index) = Trim(Me._dbcLinea_0.Text)
                If Aux <> mintVLin Then
                    If mintVLin = 0 Then
                        mblnFueraChange = True

                        'Me._dbcLinea_0.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcLinea_0.Text = cINDEFINIDA
                        mblnFueraChange = False
                    End If
                Else
                    mblnFueraChange = True
                    If mintVLin = 0 Then

                        'Me._dbcLinea_0.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcLinea_0.Text = cINDEFINIDA
                    End If
                    mblnFueraChange = False
                End If
                Call FormaDescripcion()
        End Select
    End Sub

    Private Sub _dbcLinea_0_CursorChanged(sender As Object, e As EventArgs) Handles _dbcLinea_0.CursorChanged
        Dim Index As Integer
        '= _dbcLinea_0.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String


        'cLinea(Index) = Trim(Me._dbcLinea_0.SelectedValue(Index).Text)
        cLinea(Index) = Trim(Me._dbcLinea_0.Text)
        Call FormaDescripcion()

        If mblnFueraChange Then
            Exit Sub
        End If

        Select Case Index
            Case 0 'JOYERIA

                'lStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " and descLinea = '" & Trim(Me._dbcLinea_0.SelectedValue(Index).Text) & "' Order by DescLinea "
                'ModDCombo.DCChange(lStrSql, tecla, _dbcLinea_0.SelectedValue(Index))
                lStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " and descLinea = '" & Trim(Me._dbcLinea_0.Text) & "' Order by DescLinea "
                ModDCombo.DCChange(lStrSql, tecla, _dbcLinea_0)
            Case 1 'VARIOS

                'lStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODVARIOS & " and codFamilia = " & mintVFam & " and descLinea = '" & Trim(Me._dbcLinea_0.SelectedValue(Index).Text) & "' Order by DescLinea "
                'ModDCombo.DCChange(lStrSql, tecla, _dbcLinea_0.SelectedValue(Index))
                lStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODVARIOS & " and codFamilia = " & mintVFam & " and descLinea = '" & Trim(Me._dbcLinea_0.Text) & "' Order by DescLinea "
                ModDCombo.DCChange(lStrSql, tecla, _dbcLinea_0)
        End Select


        'If Trim(Me._dbcLinea_0.SelectedValue(Index).Text) = "" Then
        If Trim(Me._dbcLinea_0.Text) = "" Then
            mintJLin = 0
            mintVLin = 0
            If Index = 0 Then
                mblnFueraChange = True
                mintJSub = 0

                Me.dbcSubLinea.Text = cINDEFINIDA
                mblnFueraChange = False
            End If
        End If
        Exit Sub
MError:
        ModEstandar.MostrarError()
    End Sub




    Private Sub _dbcLinea_1_KeyDown(sender As Object, e As KeyEventArgs) Handles _dbcLinea_1.KeyDown
        Dim Index As Integer
        '= _dbcLinea_1.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                'Me._dbcLinea_1.SelectedValue(Index).Focus()
                Me._dbcLinea_1.Focus()
            Case Else
                tecla = e.KeyCode
        End Select
    End Sub

    Private Sub _dbcLinea_1_MouseUp(sender As Object, e As MouseEventArgs) Handles _dbcLinea_1.MouseUp
        Dim Index As Integer
        '= _dbcLinea_0.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcLinea_1.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcLinea_1.Text)
        'If Me._dbcLinea_1.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._dbcLinea_1.SelectedItem <> 0 Then
        '_dbcLinea_1_Leave(_dbcLinea_1.SelectedValue.Item(Index), New System.EventArgs())
        '_dbcLinea_1_Leave(_dbcLinea_1, New System.EventArgs())
        'End If

        'Me._dbcLinea_1.SelectedValue(Index).Text = Aux
        Me._dbcLinea_1.Text = Aux
    End Sub

    Private Sub _dbcLinea_1_Enter(sender As Object, e As EventArgs) Handles _dbcLinea_1.Enter
        Dim Index As Integer
        '= _dbcLinea_1.SelectedValue(sender)
        Pon_Tool()
        Select Case Index
            Case 0 'JOYERIA
                gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " ORDER BY descLinea "
            Case 1 'VARIOS
                gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODVARIOS & " and codFamilia = " & mintVFam & " ORDER BY descLinea "
        End Select
        'ModDCombo.DCGotFocus(gStrSql, _dbcLinea_1.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _dbcLinea_1)
    End Sub

    Private Sub _dbcLinea_1_Leave(sender As Object, e As EventArgs) Handles _dbcLinea_1.Leave
        Dim Index As Integer
        '= _dbcLinea_1.SelectedValue(sender)
        Dim I As Integer
        Dim Aux As Integer
        Dim cDescripcion As String
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        'cDescripcion = Trim(Me._dbcLinea_1.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._dbcLinea_1.Text)
        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
        Select Case Index
            Case 0 'JOYERIA
                gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " and descLinea = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
                Aux = mintJLin
                mintJLin = 0
                'ModDCombo.DCLostFocus(_dbcLinea_1.SelectedValue(Index), gStrSql, mintJLin)
                ModDCombo.DCLostFocus(_dbcLinea_1, gStrSql, mintJLin)

                'cLinea(Index) = Trim(Me._dbcLinea_1.SelectedValue(Index).Text)
                cLinea(Index) = Trim(Me._dbcLinea_1.Text)
                If mintJLin <> Aux Then
                    mblnFueraChange = True
                    If mintJLin = 0 Then

                        'Me._dbcLinea_1.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcLinea_1.Text = cINDEFINIDA
                    End If
                    mintJSub = 0

                    Me.dbcSubLinea.Text = cINDEFINIDA
                    mblnFueraChange = False
                Else
                    mblnFueraChange = True
                    If mintJLin = 0 Then

                        'Me._dbcLinea_1.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcLinea_1.Text = cINDEFINIDA
                    End If
                    mblnFueraChange = False
                End If
                Call FormaDescripcion()
            Case 1 'VARIOS
                gStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODVARIOS & " and codFamilia = " & mintVFam & " and descLinea = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
                Aux = mintVLin
                mintVLin = 0
                'ModDCombo.DCLostFocus(_dbcLinea_1.SelectedValue(Index), gStrSql, mintVLin)
                ModDCombo.DCLostFocus(_dbcLinea_1, gStrSql, mintVLin)

                'cLinea(Index) = Trim(Me._dbcLinea_1.SelectedValue(Index).Text)
                cLinea(Index) = Trim(Me._dbcLinea_1.Text)
                If Aux <> mintVLin Then
                    If mintVLin = 0 Then
                        mblnFueraChange = True

                        'Me._dbcLinea_1.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcLinea_1.Text = cINDEFINIDA
                        mblnFueraChange = False
                    End If
                Else
                    mblnFueraChange = True
                    If mintVLin = 0 Then

                        'Me._dbcLinea_1.SelectedValue(Index).Text = cINDEFINIDA
                        Me._dbcLinea_1.Text = cINDEFINIDA
                    End If
                    mblnFueraChange = False
                End If
                Call FormaDescripcion()
        End Select
    End Sub

    Private Sub _dbcLinea_1_CursorChanged(sender As Object, e As EventArgs) Handles _dbcLinea_1.CursorChanged
        Dim Index As Integer
        '= _dbcLinea_1.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String


        'cLinea(Index) = Trim(Me._dbcLinea_1.SelectedValue(Index).Text)
        cLinea(Index) = Trim(Me._dbcLinea_1.Text)
        Call FormaDescripcion()

        If mblnFueraChange Then
            Exit Sub
        End If

        Select Case Index
            Case 0 'JOYERIA

                'lStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " and descLinea = '" & Trim(Me._dbcLinea_1.SelectedValue(Index).Text) & "' Order by DescLinea "
                'ModDCombo.DCChange(lStrSql, tecla, _dbcLinea_1.SelectedValue(Index))
                lStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODJOYERIA & " and codFamilia = " & mintJFam & " and descLinea = '" & Trim(Me._dbcLinea_1.Text) & "' Order by DescLinea "
                ModDCombo.DCChange(lStrSql, tecla, _dbcLinea_1)
            Case 1 'VARIOS

                'lStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODVARIOS & " and codFamilia = " & mintVFam & " and descLinea = '" & Trim(Me._dbcLinea_1.SelectedValue(Index).Text) & "' Order by DescLinea "
                'ModDCombo.DCChange(lStrSql, tecla, _dbcLinea_1.SelectedValue(Index))
                lStrSql = "SELECT codLinea, LTrim(RTrim(descLinea)) as descLinea FROM catLineas Where codGrupo = " & gCODVARIOS & " and codFamilia = " & mintVFam & " and descLinea = '" & Trim(Me._dbcLinea_1.Text) & "' Order by DescLinea "
                ModDCombo.DCChange(lStrSql, tecla, _dbcLinea_1)
        End Select


        'If Trim(Me._dbcLinea_1.SelectedValue(Index).Text) = "" Then
        If Trim(Me._dbcLinea_1.Text) = "" Then
            mintJLin = 0
            mintVLin = 0
            If Index = 0 Then
                mblnFueraChange = True
                mintJSub = 0

                Me.dbcSubLinea.Text = cINDEFINIDA
                mblnFueraChange = False
            End If
        End If
        Exit Sub
MError:
        ModEstandar.MostrarError()
    End Sub



    '    Private Sub dbcMaterial_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcMaterial.CursorChanged
    '        Dim Index As Integer = dbcMaterial.SelectedValue(eventSender)
    '        On Error GoTo MError
    '        Dim lStrSql As String
    '        Dim mintMaterial As Integer

    '        If mblnFueraChange Then Exit Sub

    '        lStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial LIKE '" & Trim(Me.dbcMaterial.SelectedValue(Index).Text) & "%' Order by DescTipoMaterial "
    '        Select Case Index
    '            Case 0 'JOYERIA
    '                ModDCombo.DCChange(lStrSql, tecla, dbcMaterial.SelectedValue(Index))
    '                '''mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(dbcMaterial(Index).text))
    '            Case 1 'RELOJERIA
    '                ModDCombo.DCChange(lStrSql, tecla, dbcMaterial.SelectedValue(Index))
    '                '''mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(dbcMaterial(Index).text))
    '            Case 2 'VARIOS
    '                ModDCombo.DCChange(lStrSql, tecla, dbcMaterial.SelectedValue(Index))
    '                '''mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(dbcMaterial(Index).text))
    '        End Select

    '        mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(dbcMaterial.SelectedValue(Index).Text))

    '        cTipoMaterial = dbcMaterial.SelectedValue(Index).Text
    '        If cTipoMaterial = "" Then
    '            cTipoMaterialDescCorta = ""
    '        Else
    '            cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintMaterial)
    '        End If

    '        Select Case Index
    '            Case 0 'JOYERIA
    '                mintJMaterial = mintMaterial
    '            Case 1 'RELOJERIA
    '                mintRMaterial = mintMaterial
    '            Case 2 'VARIOS
    '                mintVMaterial = mintMaterial
    '        End Select
    '        FormaDescripcion()


    '        If Trim(dbcMaterial.SelectedValue(Index).Text) = "" Then
    '            mintJMaterial = 0
    '            mintRMaterial = 0
    '            mintVMaterial = 0
    '        End If
    '        Exit Sub
    'MError:
    '        ModEstandar.MostrarError()
    '    End Sub

    '    Private Sub dbcMaterial_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMaterial.Enter
    '        Dim Index As Integer = dbcMaterial.SelectedValue(eventSender)
    '        Pon_Tool()
    '        gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial ORDER BY descTipoMaterial "
    '        ModDCombo.DCGotFocus(gStrSql, dbcMaterial.SelectedValue(Index))
    '    End Sub

    '    Private Sub dbcMaterial_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcMaterial.KeyDown
    '        Dim Index As Integer = dbcMaterial.SelectedValue(eventSender)
    '        Select Case eventArgs.KeyCode
    '            Case System.Windows.Forms.Keys.Escape
    '                Select Case Index
    '                    Case 0 'JOYERIA
    '                        Me.dbcKilates.Focus()
    '                        '''Me.txtCostoReal(Index).SetFocus
    '                    Case 1 'RELOJERIA
    '                        If Not _optMovimiento_0.Checked And Not _optMovimiento_1.Checked And Not _optMovimiento_2.Checked Then
    '                            dbcModelo.Focus()
    '                        ElseIf _optMovimiento_2.Checked Then
    '                            _optMovimiento_2.Focus()
    '                        ElseIf _optMovimiento_1.Checked Then
    '                            _optMovimiento_1.Focus()
    '                        ElseIf _optMovimiento_0.Checked Then
    '                            _optMovimiento_0.Focus()
    '                        End If
    '                    Case 2 'VARIOS
    '                        Me.dbcLinea.SelectedValue(1).Focus() '''06AGO2007 - MAVF
    '                End Select
    '            Case Else
    '                tecla = eventArgs.KeyCode
    '        End Select
    '    End Sub

    '    Private Sub dbcMaterial_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcMaterial.KeyUp
    '        Dim Index As Integer = dbcMaterial.SelectedValue(eventSender)
    '        Dim Aux As String

    '        Aux = Trim(Me.dbcMaterial.SelectedValue(Index).Text)
    '        If Me.dbcMaterial.SelectedValue(Index).SelectedItem <> 0 Then
    '            dbcMaterial_Leave(dbcMaterial.SelectedValue.Item(Index), New System.EventArgs())
    '        End If

    '        Me.dbcMaterial.SelectedValue(Index).Text = Aux
    '    End Sub

    '    Private Sub dbcMaterial_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcMaterial.Leave
    '        Dim Index As Integer = dbcMaterial.SelectedValue(eventSender)
    '        Dim I As Integer
    '        Dim Aux As Integer
    '        Dim cDescripcion As String

    '        cDescripcion = Trim(Me.dbcMaterial.SelectedValue(Index).Text)
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
    '            Exit Sub
    '        End If
    '        Select Case Index
    '            Case 0 'JOYERIA
    '                gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
    '                mintJMaterial = 0
    '                ModDCombo.DCLostFocus(dbcMaterial.SelectedValue(Index), gStrSql, mintJMaterial)
    '                If mintJMaterial = 0 Then
    '                    mblnFueraChange = True

    '                    dbcMaterial.SelectedValue(Index).Text = cINDEFINIDO
    '                    mblnFueraChange = False
    '                    cTipoMaterialDescCorta = ""
    '                Else
    '                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintJMaterial)
    '                End If
    '                FormaDescripcion()

    '            Case 1 'RELOJERIA
    '                gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
    '                Aux = mintRMaterial
    '                mintRMaterial = 0
    '                ModDCombo.DCLostFocus(dbcMaterial.SelectedValue(Index), gStrSql, mintRMaterial)
    '                If mintRMaterial = 0 Then
    '                    mblnFueraChange = True

    '                    Me.dbcMaterial.SelectedValue(Index).Text = cINDEFINIDO
    '                    mblnFueraChange = False
    '                    cTipoMaterialDescCorta = ""
    '                Else
    '                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintRMaterial)
    '                End If
    '                FormaDescripcion()

    '            Case 2 'VARIOS
    '                gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
    '                Aux = mintVMaterial
    '                mintVMaterial = 0
    '                ModDCombo.DCLostFocus(dbcMaterial.SelectedValue(Index), gStrSql, mintVMaterial)
    '                If mintVMaterial = 0 Then
    '                    mblnFueraChange = True

    '                    Me.dbcMaterial.SelectedValue(Index).Text = cINDEFINIDO
    '                    mblnFueraChange = False
    '                    cTipoMaterialDescCorta = ""
    '                Else
    '                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintVMaterial)
    '                End If
    '                FormaDescripcion()
    '        End Select
    '    End Sub

    '    Private Sub dbcMaterial_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcMaterial.MouseUp
    '        Dim Index As Integer = dbcMaterial.SelectedValue(eventSender)
    '        Dim Aux As String

    '        Aux = Trim(Me.dbcMaterial.SelectedValue(Index).Text)
    '        If Me.dbcMaterial.SelectedValue(Index).SelectedItem <> 0 Then
    '            dbcMaterial_Leave(dbcMaterial.SelectedValue.Item(Index), New System.EventArgs())
    '        End If

    '        Me.dbcMaterial.SelectedValue(Index).Text = Aux
    '    End Sub


    Private Sub _dbcMaterial_0_MouseUp(sender As Object, e As MouseEventArgs) Handles _dbcMaterial_0.MouseUp
        Dim Index As Integer
        '= _dbcMaterial_0.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcMaterial_0.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcMaterial_0.Text)
        'If Me._dbcMaterial_0.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._dbcMaterial_0.SelectedItem <> 0 Then
        '    '_dbcMaterial_0_Leave(_dbcMaterial_0.SelectedValue.Item(Index), New System.EventArgs())
        '    _dbcMaterial_0_Leave(_dbcMaterial_0.SelectedValue, New System.EventArgs())
        'End If

        'Me._dbcMaterial_0.SelectedValue(Index).Text = Aux
        Me._dbcMaterial_0.Text = Aux
    End Sub

    Private Sub _dbcMaterial_0_KeyDown(sender As Object, e As KeyEventArgs) Handles _dbcMaterial_0.KeyDown
        Dim Index As Integer
        '= _dbcMaterial_0.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Select Case Index
                    Case 0 'JOYERIA
                        Me.dbcKilates.Focus()
                        '''Me.txtCostoReal(Index).SetFocus
                    Case 1 'RELOJERIA
                        If Not _optMovimiento_0.Checked And Not _optMovimiento_1.Checked And Not _optMovimiento_2.Checked Then
                            dbcModelo.Focus()
                        ElseIf _optMovimiento_2.Checked Then
                            _optMovimiento_2.Focus()
                        ElseIf _optMovimiento_1.Checked Then
                            _optMovimiento_1.Focus()
                        ElseIf _optMovimiento_0.Checked Then
                            _optMovimiento_0.Focus()
                        End If
                    Case 2 'VARIOS
                        'Me.dbcLinea.SelectedValue(1).Focus() '''06AGO2007 - MAVF
                        Me._dbcLinea_0.Focus() '''06AGO2007 - MAVF
                End Select
            Case Else
                tecla = e.KeyCode
        End Select
    End Sub

    Private Sub _dbcMaterial_0_CursorChanged(sender As Object, e As EventArgs) Handles _dbcMaterial_0.CursorChanged
        Dim Index As Integer
        '= _dbcMaterial_0.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String
        Dim mintMaterial As Integer

        If mblnFueraChange Then Exit Sub

        'lStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial LIKE '" & Trim(Me.dbcMaterial.SelectedValue(Index).Text) & "%' Order by DescTipoMaterial "
        lStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial LIKE '" & Trim(Me._dbcMaterial_0.Text) & "%' Order by DescTipoMaterial "

        Select Case Index
            Case 0 'JOYERIA
                'ModDCombo.DCChange(lStrSql, tecla, dbcMaterial.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_0)
                '''mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(dbcMaterial(Index).text))
            Case 1 'RELOJERIA
                'ModDCombo.DCChange(lStrSql, tecla, dbcMaterial.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_0)
                '''mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(dbcMaterial(Index).text))
            Case 2 'VARIOS
                'ModDCombo.DCChange(lStrSql, tecla, dbcMaterial.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_0)
                '''mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(dbcMaterial(Index).text))
        End Select

        'mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(dbcMaterial.SelectedValue(Index).Text))
        mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(_dbcMaterial_0.Text))

        'cTipoMaterial = dbcMaterial.SelectedValue(Index).Text
        cTipoMaterial = _dbcMaterial_0.Text

        If cTipoMaterial = "" Then
            cTipoMaterialDescCorta = ""
        Else
            cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintMaterial)
        End If

        Select Case Index
            Case 0 'JOYERIA
                mintJMaterial = mintMaterial
            Case 1 'RELOJERIA
                mintRMaterial = mintMaterial
            Case 2 'VARIOS
                mintVMaterial = mintMaterial
        End Select
        FormaDescripcion()


        'If Trim(dbcMaterial.SelectedValue(Index).Text) = "" Then
        If Trim(_dbcMaterial_0.Text) = "" Then
            mintJMaterial = 0
            mintRMaterial = 0
            mintVMaterial = 0
        End If
        Exit Sub
MError:
        ModEstandar.MostrarError()
    End Sub

    Private Sub _dbcMaterial_0_Enter(sender As Object, e As EventArgs) Handles _dbcMaterial_0.Enter
        Dim Index As Integer
        '= _dbcMaterial_0.SelectedValue(sender)
        Pon_Tool()
        gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial ORDER BY descTipoMaterial "
        'ModDCombo.DCGotFocus(gStrSql, dbcMaterial.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _dbcMaterial_0)
    End Sub

    Private Sub _dbcMaterial_0_Leave(sender As Object, e As EventArgs) Handles _dbcMaterial_0.Leave
        Dim Index As Integer
        '= _dbcMaterial_0.SelectedValue(sender)
        Dim I As Integer
        Dim Aux As Integer
        Dim cDescripcion As String

        'cDescripcion = Trim(Me._dbcMaterial_0.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._dbcMaterial_0.Text)

        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        Select Case Index
            Case 0 'JOYERIA
                gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
                mintJMaterial = 0
                'ModDCombo.DCLostFocus(_dbcMaterial_0.SelectedValue(Index), gStrSql, mintJMaterial)
                ModDCombo.DCLostFocus(_dbcMaterial_0, gStrSql, mintJMaterial)
                If mintJMaterial = 0 Then
                    mblnFueraChange = True

                    '_dbcMaterial_0.SelectedValue(Index).Text = cINDEFINIDO
                    _dbcMaterial_0.Text = cINDEFINIDO
                    mblnFueraChange = False
                    cTipoMaterialDescCorta = ""
                Else
                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintJMaterial)
                End If
                FormaDescripcion()

            Case 1 'RELOJERIA
                gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
                Aux = mintRMaterial
                mintRMaterial = 0
                'ModDCombo.DCLostFocus(_dbcMaterial_0.SelectedValue(Index), gStrSql, mintRMaterial)
                ModDCombo.DCLostFocus(_dbcMaterial_0, gStrSql, mintRMaterial)
                If mintRMaterial = 0 Then
                    mblnFueraChange = True

                    'Me._dbcMaterial_0.SelectedValue(Index).Text = cINDEFINIDO
                    Me._dbcMaterial_0.Text = cINDEFINIDO
                    mblnFueraChange = False
                    cTipoMaterialDescCorta = ""
                Else
                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintRMaterial)
                End If
                FormaDescripcion()

            Case 2 'VARIOS
                gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
                Aux = mintVMaterial
                mintVMaterial = 0
                'ModDCombo.DCLostFocus(_dbcMaterial_0.SelectedValue(Index), gStrSql, mintVMaterial)
                ModDCombo.DCLostFocus(_dbcMaterial_0, gStrSql, mintVMaterial)
                If mintVMaterial = 0 Then
                    mblnFueraChange = True

                    'Me._dbcMaterial_0.SelectedValue(Index).Text = cINDEFINIDO
                    Me._dbcMaterial_0.Text = cINDEFINIDO
                    mblnFueraChange = False
                    cTipoMaterialDescCorta = ""
                Else
                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintVMaterial)
                End If
                FormaDescripcion()
        End Select
    End Sub

    Private Sub _dbcMaterial_0_KeyUp(sender As Object, e As KeyEventArgs) Handles _dbcMaterial_0.KeyUp
        Dim Index As Integer
        '= _dbcMaterial_0.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcMaterial_0.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcMaterial_0.Text)
        'If Me._dbcMaterial_0.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._dbcMaterial_0.SelectedItem <> 0 Then
        '    'dbcMaterial_Leave(_dbcMaterial_0.SelectedValue.Item(Index), New System.EventArgs())
        '    dbcMaterial_Leave(_dbcMaterial_0.SelectedValue, New System.EventArgs())
        'End If

        'Me._dbcMaterial_0.SelectedValue(Index).Text = Aux
        Me._dbcMaterial_0.Text = Aux
    End Sub



    Private Sub _dbcMaterial_1_MouseUp(sender As Object, e As MouseEventArgs) Handles _dbcMaterial_1.MouseUp
        Dim Index As Integer
        '= _dbcMaterial_1.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcMaterial_1.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcMaterial_1.Text)
        'If Me._dbcMaterial_1.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._dbcMaterial_1.SelectedItem <> 0 Then
        '    '_dbcMaterial_1_Leave(_dbcMaterial_1.SelectedValue.Item(Index), New System.EventArgs())
        '    _dbcMaterial_1_Leave(_dbcMaterial_1.SelectedValue, New System.EventArgs())
        'End If

        'Me._dbcMaterial_1.SelectedValue(Index).Text = Aux
        Me._dbcMaterial_1.Text = Aux
    End Sub

    Private Sub _dbcMaterial_1_KeyDown(sender As Object, e As KeyEventArgs) Handles _dbcMaterial_1.KeyDown
        Dim Index As Integer
        '= _dbcMaterial_1.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Select Case Index
                    Case 0 'JOYERIA
                        Me.dbcKilates.Focus()
                        '''Me.txtCostoReal(Index).SetFocus
                    Case 1 'RELOJERIA
                        If Not _optMovimiento_0.Checked And Not _optMovimiento_1.Checked And Not _optMovimiento_2.Checked Then
                            dbcModelo.Focus()
                        ElseIf _optMovimiento_2.Checked Then
                            _optMovimiento_2.Focus()
                        ElseIf _optMovimiento_1.Checked Then
                            _optMovimiento_1.Focus()
                        ElseIf _optMovimiento_0.Checked Then
                            _optMovimiento_0.Focus()
                        End If
                    Case 2 'VARIOS
                        'Me.dbcLinea.SelectedValue(1).Focus() '''06AGO2007 - MAVF
                        Me._dbcLinea_1.Focus() '''06AGO2007 - MAVF
                End Select
            Case Else
                tecla = e.KeyCode
        End Select
    End Sub

    Private Sub _dbcMaterial_1_CursorChanged(sender As Object, e As EventArgs) Handles _dbcMaterial_1.CursorChanged
        Dim Index As Integer
        '= _dbcMaterial_1.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String
        Dim mintMaterial As Integer

        If mblnFueraChange Then Exit Sub

        'lStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial LIKE '" & Trim(Me._dbcMaterial_1.SelectedValue(Index).Text) & "%' Order by DescTipoMaterial "
        lStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial LIKE '" & Trim(Me._dbcMaterial_1.Text) & "%' Order by DescTipoMaterial "

        Select Case Index
            Case 0 'JOYERIA
                'ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_1.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_1)
                '''mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(_dbcMaterial_1(Index).text))
            Case 1 'RELOJERIA
                'ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_1.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_1)
                '''mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(_dbcMaterial_1(Index).text))
            Case 2 'VARIOS
                'ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_1.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_1)
                '''mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(_dbcMaterial_1(Index).text))
        End Select

        'mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(_dbcMaterial_1.SelectedValue(Index).Text))
        mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(_dbcMaterial_1.Text))

        'cTipoMaterial = _dbcMaterial_1.SelectedValue(Index).Text
        cTipoMaterial = _dbcMaterial_1.Text

        If cTipoMaterial = "" Then
            cTipoMaterialDescCorta = ""
        Else
            cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintMaterial)
        End If

        Select Case Index
            Case 0 'JOYERIA
                mintJMaterial = mintMaterial
            Case 1 'RELOJERIA
                mintRMaterial = mintMaterial
            Case 2 'VARIOS
                mintVMaterial = mintMaterial
        End Select
        FormaDescripcion()


        'If Trim(_dbcMaterial_1.SelectedValue(Index).Text) = "" Then
        If Trim(_dbcMaterial_1.Text) = "" Then
            mintJMaterial = 0
            mintRMaterial = 0
            mintVMaterial = 0
        End If
        Exit Sub
MError:
        ModEstandar.MostrarError()
    End Sub

    Private Sub _dbcMaterial_1_Enter(sender As Object, e As EventArgs) Handles _dbcMaterial_1.Enter
        Dim Index As Integer
        '= _dbcMaterial_1.SelectedValue(sender)
        Pon_Tool()
        gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial ORDER BY descTipoMaterial "
        'ModDCombo.DCGotFocus(gStrSql, _dbcMaterial_1.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _dbcMaterial_1)
    End Sub

    Private Sub _dbcMaterial_1_Leave(sender As Object, e As EventArgs) Handles _dbcMaterial_1.Leave
        Dim Index As Integer
        '= _dbcMaterial_1.SelectedValue(sender)
        Dim I As Integer
        Dim Aux As Integer
        Dim cDescripcion As String

        'cDescripcion = Trim(Me._dbcMaterial_1.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._dbcMaterial_1.Text)

        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        Select Case Index
            Case 0 'JOYERIA
                gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
                mintJMaterial = 0
                'ModDCombo.DCLostFocus(_dbcMaterial_1.SelectedValue(Index), gStrSql, mintJMaterial)
                ModDCombo.DCLostFocus(_dbcMaterial_1, gStrSql, mintJMaterial)
                If mintJMaterial = 0 Then
                    mblnFueraChange = True

                    '_dbcMaterial_1.SelectedValue(Index).Text = cINDEFINIDO
                    _dbcMaterial_1.Text = cINDEFINIDO
                    mblnFueraChange = False
                    cTipoMaterialDescCorta = ""
                Else
                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintJMaterial)
                End If
                FormaDescripcion()

            Case 1 'RELOJERIA
                gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
                Aux = mintRMaterial
                mintRMaterial = 0
                'ModDCombo.DCLostFocus(_dbcMaterial_1.SelectedValue(Index), gStrSql, mintRMaterial)
                ModDCombo.DCLostFocus(_dbcMaterial_1, gStrSql, mintRMaterial)
                If mintRMaterial = 0 Then
                    mblnFueraChange = True

                    'Me._dbcMaterial_1.SelectedValue(Index).Text = cINDEFINIDO
                    Me._dbcMaterial_1.Text = cINDEFINIDO
                    mblnFueraChange = False
                    cTipoMaterialDescCorta = ""
                Else
                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintRMaterial)
                End If
                FormaDescripcion()

            Case 2 'VARIOS
                gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
                Aux = mintVMaterial
                mintVMaterial = 0
                'ModDCombo.DCLostFocus(_dbcMaterial_1.SelectedValue(Index), gStrSql, mintVMaterial)
                ModDCombo.DCLostFocus(_dbcMaterial_1, gStrSql, mintVMaterial)
                If mintVMaterial = 0 Then
                    mblnFueraChange = True

                    'Me._dbcMaterial_1.SelectedValue(Index).Text = cINDEFINIDO
                    Me._dbcMaterial_1.Text = cINDEFINIDO
                    mblnFueraChange = False
                    cTipoMaterialDescCorta = ""
                Else
                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintVMaterial)
                End If
                FormaDescripcion()
        End Select
    End Sub

    Private Sub _dbcMaterial_1_KeyUp(sender As Object, e As KeyEventArgs) Handles _dbcMaterial_1.KeyUp
        Dim Index As Integer
        '= _dbcMaterial_1.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcMaterial_1.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcMaterial_1.Text)
        'If Me._dbcMaterial_1.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._dbcMaterial_1.SelectedItem <> 0 Then
        '    _dbcMaterial_1_Leave(_dbcMaterial_1.SelectedValue.Item(Index), New System.EventArgs())
        '    _dbcMaterial_1_Leave(_dbcMaterial_1.SelectedValue, New System.EventArgs())
        'End If

        'Me._dbcMaterial_1.SelectedValue(Index).Text = Aux
        Me._dbcMaterial_1.Text = Aux
    End Sub


    Private Sub _dbcMaterial_2_MouseUp(sender As Object, e As MouseEventArgs) Handles _dbcMaterial_2.MouseUp
        Dim Index As Integer
        '= _dbcMaterial_2.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcMaterial_2.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcMaterial_2.Text)
        'If Me._dbcMaterial_2.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._dbcMaterial_2.SelectedItem <> 0 Then
        '    '_dbcMaterial_2_Leave(_dbcMaterial_2.SelectedValue.Item(Index), New System.EventArgs())
        '    _dbcMaterial_2_Leave(_dbcMaterial_2.SelectedValue, New System.EventArgs())
        'End If

        'Me._dbcMaterial_2.SelectedValue(Index).Text = Aux
        Me._dbcMaterial_2.Text = Aux
    End Sub

    Private Sub _dbcMaterial_2_KeyDown(sender As Object, e As KeyEventArgs) Handles _dbcMaterial_2.KeyDown
        Dim Index As Integer
        '= _dbcMaterial_2.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Select Case Index
                    Case 0 'JOYERIA
                        Me.dbcKilates.Focus()
                        '''Me.txtCostoReal(Index).SetFocus
                    Case 1 'RELOJERIA
                        If Not _optMovimiento_0.Checked And Not _optMovimiento_1.Checked And Not _optMovimiento_2.Checked Then
                            dbcModelo.Focus()
                        ElseIf _optMovimiento_2.Checked Then
                            _optMovimiento_2.Focus()
                        ElseIf _optMovimiento_1.Checked Then
                            _optMovimiento_1.Focus()
                        ElseIf _optMovimiento_0.Checked Then
                            _optMovimiento_0.Focus()
                        End If
                    Case 2 'VARIOS
                        'Me._dbcLinea_1.SelectedValue(1).Focus() '''06AGO2007 - MAVF
                        Me._dbcLinea_1.Focus() '''06AGO2007 - MAVF
                End Select
            Case Else
                tecla = e.KeyCode
        End Select
    End Sub

    Private Sub _dbcMaterial_2_CursorChanged(sender As Object, e As EventArgs) Handles _dbcMaterial_2.CursorChanged
        Dim Index As Integer
        '= _dbcMaterial_2.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String
        Dim mintMaterial As Integer

        If mblnFueraChange Then Exit Sub

        'lStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial LIKE '" & Trim(Me._dbcMaterial_2.SelectedValue(Index).Text) & "%' Order by DescTipoMaterial "
        lStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial LIKE '" & Trim(Me._dbcMaterial_2.Text) & "%' Order by DescTipoMaterial "

        Select Case Index
            Case 0 'JOYERIA
                'ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_2.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_2)
                '''mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(_dbcMaterial_2(Index).text))
            Case 1 'RELOJERIA
                'ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_2.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_2)
                '''mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(_dbcMaterial_2(Index).text))
            Case 2 'VARIOS
                'ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_2.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _dbcMaterial_2)
                '''mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(_dbcMaterial_2(Index).text))
        End Select

        'mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(_dbcMaterial_2.SelectedValue(Index).Text))
        mintMaterial = BuscaTipoMaterialDescCorta_Nombre(Trim(_dbcMaterial_2.Text))

        'cTipoMaterial = _dbcMaterial_2.SelectedValue(Index).Text
        cTipoMaterial = _dbcMaterial_2.Text

        If cTipoMaterial = "" Then
            cTipoMaterialDescCorta = ""
        Else
            cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintMaterial)
        End If

        Select Case Index
            Case 0 'JOYERIA
                mintJMaterial = mintMaterial
            Case 1 'RELOJERIA
                mintRMaterial = mintMaterial
            Case 2 'VARIOS
                mintVMaterial = mintMaterial
        End Select
        FormaDescripcion()


        'If Trim(_dbcMaterial_2.SelectedValue(Index).Text) = "" Then
        If Trim(_dbcMaterial_2.Text) = "" Then
            mintJMaterial = 0
            mintRMaterial = 0
            mintVMaterial = 0
        End If
        Exit Sub
MError:
        ModEstandar.MostrarError()
    End Sub

    Private Sub _dbcMaterial_2_Enter(sender As Object, e As EventArgs) Handles _dbcMaterial_2.Enter
        Dim Index As Integer
        '= _dbcMaterial_2.SelectedValue(sender)
        Pon_Tool()
        gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial ORDER BY descTipoMaterial "
        'ModDCombo.DCGotFocus(gStrSql, _dbcMaterial_2.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _dbcMaterial_2)
    End Sub

    Private Sub _dbcMaterial_2_Leave(sender As Object, e As EventArgs) Handles _dbcMaterial_2.Leave
        Dim Index As Integer
        '= _dbcMaterial_2.SelectedValue(sender)
        Dim I As Integer
        Dim Aux As Integer
        Dim cDescripcion As String

        'cDescripcion = Trim(Me._dbcMaterial_2.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._dbcMaterial_2.Text)

        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        Select Case Index
            Case 0 'JOYERIA
                gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
                mintJMaterial = 0
                'ModDCombo.DCLostFocus(_dbcMaterial_2.SelectedValue(Index), gStrSql, mintJMaterial)
                ModDCombo.DCLostFocus(_dbcMaterial_2, gStrSql, mintJMaterial)
                If mintJMaterial = 0 Then
                    mblnFueraChange = True

                    '_dbcMaterial_2.SelectedValue(Index).Text = cINDEFINIDO
                    _dbcMaterial_2.Text = cINDEFINIDO
                    mblnFueraChange = False
                    cTipoMaterialDescCorta = ""
                Else
                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintJMaterial)
                End If
                FormaDescripcion()

            Case 1 'RELOJERIA
                gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
                Aux = mintRMaterial
                mintRMaterial = 0
                'ModDCombo.DCLostFocus(_dbcMaterial_2.SelectedValue(Index), gStrSql, mintRMaterial)
                ModDCombo.DCLostFocus(_dbcMaterial_2, gStrSql, mintRMaterial)
                If mintRMaterial = 0 Then
                    mblnFueraChange = True

                    'Me._dbcMaterial_2.SelectedValue(Index).Text = cINDEFINIDO
                    Me._dbcMaterial_2.Text = cINDEFINIDO
                    mblnFueraChange = False
                    cTipoMaterialDescCorta = ""
                Else
                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintRMaterial)
                End If
                FormaDescripcion()

            Case 2 'VARIOS
                gStrSql = "SELECT codTipoMaterial, LTrim(RTrim(descTipoMaterial)) as descTipoMaterial FROM catTipoMaterial Where descTipoMaterial = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
                Aux = mintVMaterial
                mintVMaterial = 0
                'ModDCombo.DCLostFocus(_dbcMaterial_2.SelectedValue(Index), gStrSql, mintVMaterial)
                ModDCombo.DCLostFocus(_dbcMaterial_2, gStrSql, mintVMaterial)
                If mintVMaterial = 0 Then
                    mblnFueraChange = True

                    'Me._dbcMaterial_2.SelectedValue(Index).Text = cINDEFINIDO
                    Me._dbcMaterial_2.Text = cINDEFINIDO
                    mblnFueraChange = False
                    cTipoMaterialDescCorta = ""
                Else
                    cTipoMaterialDescCorta = BuscaTipoMaterialDescCorta(mintVMaterial)
                End If
                FormaDescripcion()
        End Select
    End Sub

    Private Sub _dbcMaterial_2_KeyUp(sender As Object, e As KeyEventArgs) Handles _dbcMaterial_2.KeyUp
        Dim Index As Integer
        '= _dbcMaterial_2.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcMaterial_2.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcMaterial_2.Text)
        'If Me._dbcMaterial_2.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._dbcMaterial_2.SelectedItem <> 0 Then
        '   _dbcMaterial_2_Leave(_dbcMaterial_2.SelectedValue.Item(Index), New System.EventArgs())
        '    _dbcMaterial_2_Leave(_dbcMaterial_2.SelectedValue, New System.EventArgs())
        'End If

        'Me._dbcMaterial_2.SelectedValue(Index).Text = Aux
        Me._dbcMaterial_2.Text = Aux
    End Sub





    '    Private Sub cboUnidad_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cboUnidad.CursorChanged
    '        Dim Index As Integer = cboUnidad.SelectedValue(eventSender)
    '        On Error GoTo MError
    '        Dim lStrSql As String

    '        If mblnFueraChange Then Exit Sub


    '        lStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad LIKE '" & Trim(Me.cboUnidad.SelectedValue(Index).Text) & "%' Order by descUnidad "

    '        Select Case Index
    '            Case 0 'JOYERIA
    '                ModDCombo.DCChange(lStrSql, tecla, cboUnidad.SelectedValue(Index))
    '            Case 1 'RELOJERIA
    '                ModDCombo.DCChange(lStrSql, tecla, cboUnidad.SelectedValue(Index))
    '            Case 2 'VARIOS
    '                ModDCombo.DCChange(lStrSql, tecla, cboUnidad.SelectedValue(Index))
    '        End Select


    '        If Trim(Me.cboUnidad.SelectedValue(Index).Text) = "" Then
    '            mintJUnidad = 0
    '            mintRUnidad = 0
    '            mintVUnidad = 0
    '        End If
    'MError:
    '        If Err.Number <> 0 Then
    '            ModEstandar.MostrarError()
    '        End If
    '    End Sub

    '    Private Sub cboUnidad_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboUnidad.Enter
    '        Dim Index As Integer = cboUnidad.SelectedValue(eventSender)
    '        Pon_Tool()
    '        gStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades ORDER BY descUnidad"
    '        ModDCombo.DCGotFocus(gStrSql, cboUnidad.SelectedValue(Index))
    '    End Sub

    '    Private Sub cboUnidad_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cboUnidad.KeyDown
    '        Dim Index As Integer = cboUnidad.SelectedValue(eventSender)
    '        Select Case eventArgs.KeyCode
    '            Case System.Windows.Forms.Keys.Escape
    '                Select Case Index
    '                    Case 0 'JOYERIA
    '                        Me.txtCostoIndirecto(Index).Focus()
    '                        '''Me.dbcMaterial(Index).SetFocus
    '                    Case 1 'RELOJERIA
    '                        Me.txtCostoReal(Index).Focus()
    '                    Case 2 'VARIOS
    '                        Me.dbcMaterial.SelectedValue(Index).Focus()
    '                End Select
    '                eventSender.KeyCode = 0
    '            Case Else
    '                Select Case Index
    '                    Case 0 'JOYERIA
    '                        tecla = eventArgs.KeyCode
    '                    Case 1 'RELOJERIA
    '                        tecla = eventArgs.KeyCode
    '                    Case 2 'VARIOS
    '                        tecla = eventArgs.KeyCode
    '                End Select
    '        End Select
    '    End Sub

    '    Private Sub cboUnidad_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboUnidad.Leave
    '        Dim Index As Integer = cboUnidad.SelectedValue(eventSender)
    '        Dim I As Integer
    '        Dim cDescripcion As String

    '        cDescripcion = Trim(Me.cboUnidad.SelectedValue(Index).Text)
    '        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
    '            Exit Sub
    '        End If
    '        gStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
    '        mintJUnidad = 0
    '        mintRUnidad = 0
    '        mintVUnidad = 0
    '        Select Case Me.sstArticulo.SelectedIndex
    '            Case nJOYERIA
    '                ModDCombo.DCLostFocus(cboUnidad.SelectedValue(Index), gStrSql, mintJUnidad)
    '            Case nRELOJERIA
    '                ModDCombo.DCLostFocus(cboUnidad.SelectedValue(Index), gStrSql, mintRUnidad)
    '            Case nVARIOS
    '                ModDCombo.DCLostFocus(cboUnidad.SelectedValue(Index), gStrSql, mintVUnidad)
    '        End Select
    '        If mintJUnidad = 0 And mintRUnidad = 0 And mintVUnidad = 0 Then
    '            mblnFueraChange = True

    '            Me.cboUnidad.SelectedValue(Index).Text = cINDEFINIDA
    '            mblnFueraChange = False
    '        End If
    '    End Sub

    '    Private Sub cboUnidad_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cboUnidad.MouseUp
    '        Dim Index As Integer = cboUnidad.SelectedValue(eventSender)
    '        Dim Aux As String

    '        Aux = Trim(Me.cboUnidad.SelectedValue(Index).Text)
    '        If Me.cboUnidad.SelectedValue(Index).SelectedItem <> 0 Then
    '            cboUnidad_Leave(cboUnidad.SelectedValue.Item(Index), New System.EventArgs())
    '        End If

    '        Me.cboUnidad.SelectedValue(Index).Text = Aux
    '    End Sub


    Private Sub _cboUnidad_0_CursorChanged(sender As Object, e As EventArgs) Handles _cboUnidad_0.CursorChanged
        Dim Index As Integer
        '= _cboUnidad_0.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        'lStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad LIKE '" & Trim(Me._cboUnidad_0.SelectedValue(Index).Text) & "%' Order by descUnidad "
        lStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad LIKE '" & Trim(Me._cboUnidad_0.Text) & "%' Order by descUnidad "

        Select Case Index
            Case 0 'JOYERIA
                'ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_0.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_0)
            Case 1 'RELOJERIA
                'ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_0.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_0)
            Case 2 'VARIOS
                'ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_0.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_0)
        End Select


        'If Trim(Me._cboUnidad_0.SelectedValue(Index).Text) = "" Then
        If Trim(Me._cboUnidad_0.Text) = "" Then
            mintJUnidad = 0
            mintRUnidad = 0
            mintVUnidad = 0
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub _cboUnidad_0_Enter(sender As Object, e As EventArgs) Handles _cboUnidad_0.Enter
        Dim Index As Integer
        '= _cboUnidad_0.SelectedValue(sender)
        Pon_Tool()
        gStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades ORDER BY descUnidad"
        'ModDCombo.DCGotFocus(gStrSql, _cboUnidad_0.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _cboUnidad_0)
    End Sub

    Private Sub _cboUnidad_0_Leave(sender As Object, e As EventArgs) Handles _cboUnidad_0.Leave
        Dim Index As Integer
        '= _cboUnidad_0.SelectedValue(sender)
        Dim I As Integer
        Dim cDescripcion As String

        'cDescripcion = Trim(Me._cboUnidad_0.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._cboUnidad_0.Text)

        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        gStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
        mintJUnidad = 0
        mintRUnidad = 0
        mintVUnidad = 0
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                'ModDCombo.DCLostFocus(_cboUnidad_0.SelectedValue(Index), gStrSql, mintJUnidad)
                ModDCombo.DCLostFocus(_cboUnidad_0, gStrSql, mintJUnidad)
            Case nRELOJERIA
                'ModDCombo.DCLostFocus(_cboUnidad_0.SelectedValue(Index), gStrSql, mintRUnidad)
                ModDCombo.DCLostFocus(_cboUnidad_0, gStrSql, mintRUnidad)
            Case nVARIOS
                'ModDCombo.DCLostFocus(_cboUnidad_0.SelectedValue(Index), gStrSql, mintVUnidad)
                ModDCombo.DCLostFocus(_cboUnidad_0, gStrSql, mintVUnidad)
        End Select
        If mintJUnidad = 0 And mintRUnidad = 0 And mintVUnidad = 0 Then
            mblnFueraChange = True

            'Me._cboUnidad_0.SelectedValue(Index).Text = cINDEFINIDA
            Me._cboUnidad_0.Text = cINDEFINIDA
            mblnFueraChange = False
        End If
    End Sub

    Private Sub _cboUnidad_0_MouseUp(sender As Object, e As MouseEventArgs) Handles _cboUnidad_0.MouseUp
        Dim Index As Integer
        '= _cboUnidad_0.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._cboUnidad_0.SelectedValue(Index).Text)
        Aux = Trim(Me._cboUnidad_0.Text)
        'If Me._cboUnidad_0.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._cboUnidad_0.SelectedValue <> 0 Then
        '    '_cboUnidad_0_Leave(_cboUnidad_0.SelectedValue.Item(Index), New System.EventArgs())
        '    _cboUnidad_0_Leave(_cboUnidad_0, New System.EventArgs())
        'End If

        'Me._cboUnidad_0.SelectedValue(Index).Text = Aux
        Me._cboUnidad_0.Text = Aux
    End Sub

    Private Sub _cboUnidad_0_KeyDown(sender As Object, e As KeyEventArgs) Handles _cboUnidad_0.KeyDown
        Dim Index As Integer
        '= _cboUnidad_0.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Select Case Index
                    Case 0 'JOYERIA
                        Me.txtCostoIndirecto(Index).Focus()
                        '''Me.dbcMaterial(Index).SetFocus
                    Case 1 'RELOJERIA
                        'Me.txtCostoReal(Index).Focus()
                        Me.txtCostoReal(Index).Focus()
                    Case 2 'VARIOS
                        'Me._cboUnidad_0.SelectedValue(Index).Focus()
                        Me._cboUnidad_0.Focus()
                End Select
                sender.KeyCode = 0
            Case Else
                Select Case Index
                    Case 0 'JOYERIA
                        tecla = sender.KeyCode
                    Case 1 'RELOJERIA
                        tecla = sender.KeyCode
                    Case 2 'VARIOS
                        tecla = sender.KeyCode
                End Select
        End Select
    End Sub



    Private Sub _cboUnidad_1_CursorChanged(sender As Object, e As EventArgs) Handles _cboUnidad_1.CursorChanged
        Dim Index As Integer
        '= _cboUnidad_1.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        'lStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad LIKE '" & Trim(Me._cboUnidad_1.SelectedValue(Index).Text) & "%' Order by descUnidad "
        lStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad LIKE '" & Trim(Me._cboUnidad_1.Text) & "%' Order by descUnidad "

        Select Case Index
            Case 0 'JOYERIA
                'ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_1.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_1)
            Case 1 'RELOJERIA
                'ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_1.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_1)
            Case 2 'VARIOS
                'ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_1.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_1)
        End Select


        'If Trim(Me._cboUnidad_1.SelectedValue(Index).Text) = "" Then
        If Trim(Me._cboUnidad_1.Text) = "" Then
            mintJUnidad = 0
            mintRUnidad = 0
            mintVUnidad = 0
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub _cboUnidad_1_Enter(sender As Object, e As EventArgs) Handles _cboUnidad_1.Enter
        Dim Index As Integer
        '= _cboUnidad_1.SelectedValue(sender)
        Pon_Tool()
        gStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades ORDER BY descUnidad"
        'ModDCombo.DCGotFocus(gStrSql, _cboUnidad_1.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _cboUnidad_1)
    End Sub

    Private Sub _cboUnidad_1_Leave(sender As Object, e As EventArgs) Handles _cboUnidad_1.Leave
        Dim Index As Integer
        '= _cboUnidad_1.SelectedValue(sender)
        Dim I As Integer
        Dim cDescripcion As String

        'cDescripcion = Trim(Me._cboUnidad_1.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._cboUnidad_1.Text)

        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        gStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
        mintJUnidad = 0
        mintRUnidad = 0
        mintVUnidad = 0
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                'ModDCombo.DCLostFocus(_cboUnidad_1.SelectedValue(Index), gStrSql, mintJUnidad)
                ModDCombo.DCLostFocus(_cboUnidad_1, gStrSql, mintJUnidad)
            Case nRELOJERIA
                'ModDCombo.DCLostFocus(_cboUnidad_1.SelectedValue(Index), gStrSql, mintRUnidad)
                ModDCombo.DCLostFocus(_cboUnidad_1, gStrSql, mintRUnidad)
            Case nVARIOS
                'ModDCombo.DCLostFocus(_cboUnidad_1.SelectedValue(Index), gStrSql, mintVUnidad)
                ModDCombo.DCLostFocus(_cboUnidad_1, gStrSql, mintVUnidad)
        End Select
        If mintJUnidad = 0 And mintRUnidad = 0 And mintVUnidad = 0 Then
            mblnFueraChange = True

            'Me._cboUnidad_1.SelectedValue(Index).Text = cINDEFINIDA
            Me._cboUnidad_1.Text = cINDEFINIDA
            mblnFueraChange = False
        End If
    End Sub

    Private Sub _cboUnidad_1_MouseUp(sender As Object, e As MouseEventArgs) Handles _cboUnidad_1.MouseUp
        Dim Index As Integer
        '= _cboUnidad_1.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._cboUnidad_1.SelectedValue(Index).Text)
        Aux = Trim(Me._cboUnidad_1.Text)
        'If Me._cboUnidad_1.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._cboUnidad_1.SelectedValue <> 0 Then
        '    '_cboUnidad_1_Leave(_cboUnidad_1.SelectedValue.Item(Index), New System.EventArgs())
        '    _cboUnidad_1_Leave(_cboUnidad_1, New System.EventArgs())
        'End If

        'Me._cboUnidad_1.SelectedValue(Index).Text = Aux
        Me._cboUnidad_1.Text = Aux
    End Sub

    Private Sub _cboUnidad_1_KeyDown(sender As Object, e As KeyEventArgs) Handles _cboUnidad_1.KeyDown
        Dim Index As Integer
        '= _cboUnidad_1.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Select Case Index
                    Case 0 'JOYERIA
                        Me.txtCostoIndirecto(Index).Focus()
                        '''Me.dbcMaterial(Index).SetFocus
                    Case 1 'RELOJERIA
                        'Me.txtCostoReal(Index).Focus()
                        Me.txtCostoReal(Index).Focus()
                    Case 2 'VARIOS
                        'Me._cboUnidad_1.SelectedValue(Index).Focus()
                        Me._cboUnidad_1.Focus()
                End Select
                sender.KeyCode = 0
            Case Else
                Select Case Index
                    Case 0 'JOYERIA
                        tecla = sender.KeyCode
                    Case 1 'RELOJERIA
                        tecla = sender.KeyCode
                    Case 2 'VARIOS
                        tecla = sender.KeyCode
                End Select
        End Select
    End Sub




    Private Sub _cboUnidad_2_CursorChanged(sender As Object, e As EventArgs) Handles _cboUnidad_2.CursorChanged
        Dim Index As Integer
        '= _cboUnidad_2.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        'lStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad LIKE '" & Trim(Me._cboUnidad_2.SelectedValue(Index).Text) & "%' Order by descUnidad "
        lStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad LIKE '" & Trim(Me._cboUnidad_2.Text) & "%' Order by descUnidad "

        Select Case Index
            Case 0 'JOYERIA
                'ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_2.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_2)
            Case 1 'RELOJERIA
                'ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_2.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_2)
            Case 2 'VARIOS
                'ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_2.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboUnidad_2)
        End Select


        'If Trim(Me._cboUnidad_2.SelectedValue(Index).Text) = "" Then
        If Trim(Me._cboUnidad_2.Text) = "" Then
            mintJUnidad = 0
            mintRUnidad = 0
            mintVUnidad = 0
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub _cboUnidad_2_Enter(sender As Object, e As EventArgs) Handles _cboUnidad_2.Enter
        Dim Index As Integer
        '= _cboUnidad_2.SelectedValue(sender)
        Pon_Tool()
        gStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades ORDER BY descUnidad"
        'ModDCombo.DCGotFocus(gStrSql, _cboUnidad_2.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _cboUnidad_2)
    End Sub

    Private Sub _cboUnidad_2_Leave(sender As Object, e As EventArgs) Handles _cboUnidad_2.Leave
        Dim Index As Integer
        '= _cboUnidad_2.SelectedValue(sender)
        Dim I As Integer
        Dim cDescripcion As String

        'cDescripcion = Trim(Me._cboUnidad_2.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._cboUnidad_2.Text)

        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "%'")
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        gStrSql = "SELECT codUnidad, LTrim(RTrim(descUnidad)) as descUnidad FROM catUnidades Where descUnidad = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDA), "'", Trim(cDescripcion) & "'")
        mintJUnidad = 0
        mintRUnidad = 0
        mintVUnidad = 0
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                'ModDCombo.DCLostFocus(_cboUnidad_2.SelectedValue(Index), gStrSql, mintJUnidad)
                ModDCombo.DCLostFocus(_cboUnidad_2, gStrSql, mintJUnidad)
            Case nRELOJERIA
                'ModDCombo.DCLostFocus(_cboUnidad_2.SelectedValue(Index), gStrSql, mintRUnidad)
                ModDCombo.DCLostFocus(_cboUnidad_2, gStrSql, mintRUnidad)
            Case nVARIOS
                'ModDCombo.DCLostFocus(_cboUnidad_2.SelectedValue(Index), gStrSql, mintVUnidad)
                ModDCombo.DCLostFocus(_cboUnidad_2, gStrSql, mintVUnidad)
        End Select
        If mintJUnidad = 0 And mintRUnidad = 0 And mintVUnidad = 0 Then
            mblnFueraChange = True

            'Me._cboUnidad_2.SelectedValue(Index).Text = cINDEFINIDA
            Me._cboUnidad_2.Text = cINDEFINIDA
            mblnFueraChange = False
        End If
    End Sub

    Private Sub _cboUnidad_2_MouseUp(sender As Object, e As MouseEventArgs) Handles _cboUnidad_2.MouseUp
        Dim Index As Integer
        '= _cboUnidad_2.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._cboUnidad_2.SelectedValue(Index).Text)
        Aux = Trim(Me._cboUnidad_2.Text)
        'If Me._cboUnidad_2.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._cboUnidad_2.SelectedValue <> 0 Then
        '    '_cboUnidad_2_Leave(_cboUnidad_2.SelectedValue.Item(Index), New System.EventArgs())
        '   _cboUnidad_2_Leave(_cboUnidad_2, New System.EventArgs())
        'End If

        'Me._cboUnidad_2.SelectedValue(Index).Text = Aux
        Me._cboUnidad_2.Text = Aux
    End Sub

    Private Sub _cboUnidad_2_KeyDown(sender As Object, e As KeyEventArgs) Handles _cboUnidad_2.KeyDown
        Dim Index As Integer
        '= _cboUnidad_2.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Select Case Index
                    Case 0 'JOYERIA
                        Me.txtCostoIndirecto(Index).Focus()
                        '''Me.dbcMaterial(Index).SetFocus
                    Case 1 'RELOJERIA
                        'Me.txtCostoReal(Index).Focus()
                        Me.txtCostoReal(Index).Focus()
                    Case 2 'VARIOS
                        'Me._cboUnidad_2.SelectedValue(Index).Focus()
                        Me._cboUnidad_2.Focus()
                End Select
                sender.KeyCode = 0
            Case Else
                Select Case Index
                    Case 0 'JOYERIA
                        tecla = sender.KeyCode
                    Case 1 'RELOJERIA
                        tecla = sender.KeyCode
                    Case 2 'VARIOS
                        tecla = sender.KeyCode
                End Select
        End Select
    End Sub


    '    Private Sub cboAlmacen_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboAlmacen.CursorChanged
    '        Dim Index As Integer = cboAlmacen.SelectedValue(eventSender)
    '        On Error GoTo MError
    '        Dim lStrSql As String

    '        If mblnFueraChange Then Exit Sub


    '        lStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me.cboAlmacen.SelectedValue(Index).Text) & "%' Order by CodAlmacenOrigen "

    '        Select Case Index
    '            Case 0 'JOYERIA
    '                ModDCombo.DCChange(lStrSql, tecla, cboAlmacen.SelectedValue(Index))
    '            Case 1 'RELOJERIA
    '                ModDCombo.DCChange(lStrSql, tecla, cboAlmacen.SelectedValue(Index))
    '            Case 2 'VARIOS
    '                ModDCombo.DCChange(lStrSql, tecla, cboAlmacen.SelectedValue(Index))
    '        End Select


    '        If Trim(Me.cboAlmacen.SelectedValue(Index).Text) = "" Then
    '            mintJOrigen = 0
    '            mintROrigen = 0
    '            mintVOrigen = 0
    '        End If
    'MError:
    '        If Err.Number <> 0 Then
    '            ModEstandar.MostrarError()
    '        End If
    '    End Sub

    '    Private Sub cboAlmacen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboAlmacen.Enter
    '        Dim Index As Integer = cboAlmacen.SelectedValue(eventSender)
    '        Pon_Tool()
    '        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen ORDER BY CodAlmacenOrigen "
    '        ModDCombo.DCGotFocus(gStrSql, cboAlmacen.SelectedValue(Index))
    '    End Sub

    '    Private Sub cboAlmacen_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cboAlmacen.KeyDown
    '        Dim Index As Integer = cboAlmacen.SelectedValue(eventSender)
    '        Select Case eventArgs.KeyCode
    '            Case System.Windows.Forms.Keys.Escape
    '                Select Case Index
    '                    Case 0 'JOYERIA
    '                        Me.cboUnidad.SelectedValue(Index).Focus()
    '                    Case 1 'RELOJERIA
    '                        Me.cboUnidad.SelectedValue(Index).Focus()
    '                    Case 2 'VARIOS
    '                        Me.cboUnidad.SelectedValue(Index).Focus()
    '                End Select
    '            Case Else
    '                Select Case Index
    '                    Case 0 'JOYERIA
    '                        tecla = eventArgs.KeyCode
    '                    Case 1 'RELOJERIA
    '                        tecla = eventArgs.KeyCode
    '                    Case 2 'VARIOS
    '                        tecla = eventArgs.KeyCode
    '                End Select
    '        End Select
    '    End Sub

    '    Private Sub cboAlmacen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboAlmacen.Leave
    '        Dim Index As Integer = cboAlmacen.SelectedValue(eventSender)
    '        Dim I As Integer
    '        Dim cDescripcion As String
    '        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
    '            Exit Sub
    '        End If

    '        cDescripcion = Trim(Me.cboAlmacen.SelectedValue(Index).Text)
    '        If Trim(cDescripcion) = Trim(cINDEFINIDO) Then
    '            cDescripcion = ""
    '        End If
    '        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "%'")
    '        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen = '" & IIf(Trim(cDescripcion) = "", "'", Trim(cDescripcion) & "'")
    '        mintJOrigen = 0
    '        mintROrigen = 0
    '        mintVOrigen = 0
    '        Select Case Me.sstArticulo.SelectedIndex
    '            Case nJOYERIA
    '                ModDCombo.DCLostFocus(cboAlmacen.SelectedValue(Index), gStrSql, mintJOrigen)
    '            Case nRELOJERIA
    '                ModDCombo.DCLostFocus(cboAlmacen.SelectedValue(Index), gStrSql, mintROrigen)
    '            Case nVARIOS
    '                ModDCombo.DCLostFocus(cboAlmacen.SelectedValue(Index), gStrSql, mintVOrigen)
    '        End Select
    '        If mintJOrigen = 0 And mintROrigen = 0 And mintVOrigen = 0 And cDescripcion = "" Then
    '            mblnFueraChange = True

    '            Me.cboAlmacen.SelectedValue(Index).Text = cINDEFINIDO
    '            mblnFueraChange = False
    '        End If
    '    End Sub

    '    Private Sub cboAlmacen_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cboAlmacen.MouseUp
    '        Dim Index As Integer = cboAlmacen.SelectedValue(eventSender)
    '        Dim Aux As String

    '        Aux = Trim(Me.cboAlmacen.SelectedValue(Index).Text)
    '        If Me.cboAlmacen.SelectedValue(Index).SelectedItem <> 0 Then
    '            cboAlmacen_Leave(cboAlmacen.SelectedValue.Item(Index), New System.EventArgs())
    '        End If

    '        Me.cboAlmacen.SelectedValue(Index).Text = Aux
    '    End Sub


    Private Sub _cboAlmacen_0_Enter(sender As Object, e As EventArgs) Handles _cboAlmacen_0.Enter
        Dim Index As Integer
        '= _cboAlmacen_0.SelectedValue(sender)
        Pon_Tool()
        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen ORDER BY CodAlmacenOrigen "
        'ModDCombo.DCGotFocus(gStrSql, _cboAlmacen_0.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _cboAlmacen_0)
    End Sub

    Private Sub _cboAlmacen_0_Leave(sender As Object, e As EventArgs) Handles _cboAlmacen_0.Leave
        Dim Index As Integer
        '= _cboAlmacen_0.SelectedValue(sender)
        Dim I As Integer
        Dim cDescripcion As String
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        'cDescripcion = Trim(Me._cboAlmacen_0.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._cboAlmacen_0.Text)
        If Trim(cDescripcion) = Trim(cINDEFINIDO) Then
            cDescripcion = ""
        End If
        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "%'")
        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen = '" & IIf(Trim(cDescripcion) = "", "'", Trim(cDescripcion) & "'")
        mintJOrigen = 0
        mintROrigen = 0
        mintVOrigen = 0
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                'ModDCombo.DCLostFocus(_cboAlmacen_0.SelectedValue(Index), gStrSql, mintJOrigen)
                ModDCombo.DCLostFocus(_cboAlmacen_0, gStrSql, mintJOrigen)
            Case nRELOJERIA
                'ModDCombo.DCLostFocus(_cboAlmacen_0.SelectedValue(Index), gStrSql, mintROrigen)
                ModDCombo.DCLostFocus(_cboAlmacen_0, gStrSql, mintROrigen)
            Case nVARIOS
                'ModDCombo.DCLostFocus(_cboAlmacen_0.SelectedValue(Index), gStrSql, mintVOrigen)
                ModDCombo.DCLostFocus(_cboAlmacen_0, gStrSql, mintVOrigen)
        End Select
        If mintJOrigen = 0 And mintROrigen = 0 And mintVOrigen = 0 And cDescripcion = "" Then
            mblnFueraChange = True

            'Me._cboAlmacen_0.SelectedValue(Index).Text = cINDEFINIDO
            Me._cboAlmacen_0.Text = cINDEFINIDO
            mblnFueraChange = False
        End If
    End Sub

    Private Sub _cboAlmacen_0_KeyDown(sender As Object, e As KeyEventArgs) Handles _cboAlmacen_0.KeyDown
        Dim Index As Integer
        '= _cboAlmacen_0.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Select Case Index
                    Case 0 'JOYERIA
                        'Me._cboAlmacen_0.SelectedValue(Index).Focus()
                        Me._cboAlmacen_0.Focus()
                    Case 1 'RELOJERIA
                        'Me._cboAlmacen_0.SelectedValue(Index).Focus()
                        Me._cboAlmacen_0.Focus()
                    Case 2 'VARIOS
                        'Me._cboAlmacen_0.SelectedValue(Index).Focus()
                        Me._cboAlmacen_0.Focus()
                End Select
            Case Else
                Select Case Index
                    Case 0 'JOYERIA
                        tecla = sender.KeyCode
                    Case 1 'RELOJERIA
                        tecla = sender.KeyCode
                    Case 2 'VARIOS
                        tecla = sender.KeyCode
                End Select
        End Select
    End Sub

    Private Sub _cboAlmacen_0_MouseUp(sender As Object, e As MouseEventArgs) Handles _cboAlmacen_0.MouseUp
        Dim Index As Integer
        '= _cboAlmacen_0.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._cboAlmacen_0.SelectedValue(Index).Text)
        Aux = Trim(Me._cboAlmacen_0.Text)
        'If Me._cboAlmacen_0.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._cboAlmacen_0.SelectedValue <> 0 Then
        '    '_cboAlmacen_0_Leave(_cboAlmacen_0.SelectedValue.Item(Index), New System.EventArgs())
        '    _cboAlmacen_0_Leave(_cboAlmacen_0.SelectedValue, New System.EventArgs())
        'End If

        'Me._cboAlmacen_0.SelectedValue(Index).Text = Aux
        Me._cboAlmacen_0.Text = Aux
    End Sub

    Private Sub _cboAlmacen_0_CursorChanged(sender As Object, e As EventArgs) Handles _cboAlmacen_0.CursorChanged
        Dim Index As Integer
        '= _cboAlmacen_0.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        'lStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me._cboAlmacen_0.SelectedValue(Index).Text) & "%' Order by CodAlmacenOrigen "
        lStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me._cboAlmacen_0.Text) & "%' Order by CodAlmacenOrigen "

        Select Case Index
            Case 0 'JOYERIA
                'ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_0.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_0)
            Case 1 'RELOJERIA
                'ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_0.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_0)
            Case 2 'VARIOS
                'ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_0.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_0)
        End Select


        'If Trim(Me._cboAlmacen_0.SelectedValue(Index).Text) = "" Then
        If Trim(Me._cboAlmacen_0.Text) = "" Then
            mintJOrigen = 0
            mintROrigen = 0
            mintVOrigen = 0
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub


    Private Sub _cboAlmacen_1_Enter(sender As Object, e As EventArgs) Handles _cboAlmacen_1.Enter
        Dim Index As Integer
        '= _cboAlmacen_1.SelectedValue(sender)
        Pon_Tool()
        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen ORDER BY CodAlmacenOrigen "
        'ModDCombo.DCGotFocus(gStrSql, _cboAlmacen_1.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _cboAlmacen_1)
    End Sub

    Private Sub _cboAlmacen_1_Leave(sender As Object, e As EventArgs) Handles _cboAlmacen_1.Leave
        Dim Index As Integer
        '= _cboAlmacen_1.SelectedValue(sender)
        Dim I As Integer
        Dim cDescripcion As String
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        'cDescripcion = Trim(Me._cboAlmacen_1.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._cboAlmacen_1.Text)
        If Trim(cDescripcion) = Trim(cINDEFINIDO) Then
            cDescripcion = ""
        End If
        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "%'")
        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen = '" & IIf(Trim(cDescripcion) = "", "'", Trim(cDescripcion) & "'")
        mintJOrigen = 0
        mintROrigen = 0
        mintVOrigen = 0
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                'ModDCombo.DCLostFocus(_cboAlmacen_1.SelectedValue(Index), gStrSql, mintJOrigen)
                ModDCombo.DCLostFocus(_cboAlmacen_1, gStrSql, mintJOrigen)
            Case nRELOJERIA
                'ModDCombo.DCLostFocus(_cboAlmacen_1.SelectedValue(Index), gStrSql, mintROrigen)
                ModDCombo.DCLostFocus(_cboAlmacen_1, gStrSql, mintROrigen)
            Case nVARIOS
                'ModDCombo.DCLostFocus(_cboAlmacen_1.SelectedValue(Index), gStrSql, mintVOrigen)
                ModDCombo.DCLostFocus(_cboAlmacen_1, gStrSql, mintVOrigen)
        End Select
        If mintJOrigen = 0 And mintROrigen = 0 And mintVOrigen = 0 And cDescripcion = "" Then
            mblnFueraChange = True

            'Me._cboAlmacen_1.SelectedValue(Index).Text = cINDEFINIDO
            Me._cboAlmacen_1.Text = cINDEFINIDO
            mblnFueraChange = False
        End If
    End Sub

    Private Sub _cboAlmacen_1_KeyDown(sender As Object, e As KeyEventArgs) Handles _cboAlmacen_1.KeyDown
        Dim Index As Integer
        '= _cboAlmacen_1.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Select Case Index
                    Case 0 'JOYERIA
                        'Me._cboAlmacen_1.SelectedValue(Index).Focus()
                        Me._cboAlmacen_1.Focus()
                    Case 1 'RELOJERIA
                        'Me._cboAlmacen_1.SelectedValue(Index).Focus()
                        Me._cboAlmacen_1.Focus()
                    Case 2 'VARIOS
                        'Me._cboAlmacen_1.SelectedValue(Index).Focus()
                        Me._cboAlmacen_1.Focus()
                End Select
            Case Else
                Select Case Index
                    Case 0 'JOYERIA
                        tecla = sender.KeyCode
                    Case 1 'RELOJERIA
                        tecla = sender.KeyCode
                    Case 2 'VARIOS
                        tecla = sender.KeyCode
                End Select
        End Select
    End Sub

    Private Sub _cboAlmacen_1_MouseUp(sender As Object, e As MouseEventArgs) Handles _cboAlmacen_1.MouseUp
        Dim Index As Integer
        '= _cboAlmacen_1.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._cboAlmacen_1.SelectedValue(Index).Text)
        Aux = Trim(Me._cboAlmacen_1.Text)
        'If Me._cboAlmacen_1.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._cboAlmacen_1.SelectedValue <> 0 Then
        '    '_cboAlmacen_1_Leave(_cboAlmacen_1.SelectedValue.Item(Index), New System.EventArgs())
        '    _cboAlmacen_1_Leave(_cboAlmacen_1.SelectedValue, New System.EventArgs())
        'End If

        'Me._cboAlmacen_1.SelectedValue(Index).Text = Aux
        Me._cboAlmacen_1.Text = Aux
    End Sub

    Private Sub _cboAlmacen_1_CursorChanged(sender As Object, e As EventArgs) Handles _cboAlmacen_1.CursorChanged
        Dim Index As Integer
        '= _cboAlmacen_1.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        'lStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me._cboAlmacen_1.SelectedValue(Index).Text) & "%' Order by CodAlmacenOrigen "
        lStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me._cboAlmacen_1.Text) & "%' Order by CodAlmacenOrigen "

        Select Case Index
            Case 0 'JOYERIA
                'ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_1.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_1)
            Case 1 'RELOJERIA
                'ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_1.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_1)
            Case 2 'VARIOS
                'ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_1.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_1)
        End Select


        'If Trim(Me._cboAlmacen_1.SelectedValue(Index).Text) = "" Then
        If Trim(Me._cboAlmacen_1.Text) = "" Then
            mintJOrigen = 0
            mintROrigen = 0
            mintVOrigen = 0
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub


    Private Sub _cboAlmacen_2_Enter(sender As Object, e As EventArgs) Handles _cboAlmacen_2.Enter
        Dim Index As Integer
        '= _cboAlmacen_2.SelectedValue(sender)
        Pon_Tool()
        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen ORDER BY CodAlmacenOrigen "
        'ModDCombo.DCGotFocus(gStrSql, _cboAlmacen_2.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _cboAlmacen_2)
    End Sub

    Private Sub _cboAlmacen_2_Leave(sender As Object, e As EventArgs) Handles _cboAlmacen_2.Leave
        Dim Index As Integer
        '= _cboAlmacen_2.SelectedValue(sender)
        Dim I As Integer
        Dim cDescripcion As String
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If

        'cDescripcion = Trim(Me._cboAlmacen_2.SelectedValue(Index).Text)
        cDescripcion = Trim(Me._cboAlmacen_2.Text)
        If Trim(cDescripcion) = Trim(cINDEFINIDO) Then
            cDescripcion = ""
        End If
        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "%'")
        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen = '" & IIf(Trim(cDescripcion) = "", "'", Trim(cDescripcion) & "'")
        mintJOrigen = 0
        mintROrigen = 0
        mintVOrigen = 0
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                'ModDCombo.DCLostFocus(_cboAlmacen_2.SelectedValue(Index), gStrSql, mintJOrigen)
                ModDCombo.DCLostFocus(_cboAlmacen_2, gStrSql, mintJOrigen)
            Case nRELOJERIA
                'ModDCombo.DCLostFocus(_cboAlmacen_2.SelectedValue(Index), gStrSql, mintROrigen)
                ModDCombo.DCLostFocus(_cboAlmacen_2, gStrSql, mintROrigen)
            Case nVARIOS
                'ModDCombo.DCLostFocus(_cboAlmacen_2.SelectedValue(Index), gStrSql, mintVOrigen)
                ModDCombo.DCLostFocus(_cboAlmacen_2, gStrSql, mintVOrigen)
        End Select
        If mintJOrigen = 0 And mintROrigen = 0 And mintVOrigen = 0 And cDescripcion = "" Then
            mblnFueraChange = True

            'Me._cboAlmacen_2.SelectedValue(Index).Text = cINDEFINIDO
            Me._cboAlmacen_2.Text = cINDEFINIDO
            mblnFueraChange = False
        End If
    End Sub

    Private Sub _cboAlmacen_2_KeyDown(sender As Object, e As KeyEventArgs) Handles _cboAlmacen_2.KeyDown
        Dim Index As Integer
        '= _cboAlmacen_2.SelectedValue(sender)
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Escape
                Select Case Index
                    Case 0 'JOYERIA
                        'Me._cboAlmacen_2.SelectedValue(Index).Focus()
                        Me._cboAlmacen_2.Focus()
                    Case 1 'RELOJERIA
                        'Me._cboAlmacen_2.SelectedValue(Index).Focus()
                        Me._cboAlmacen_2.Focus()
                    Case 2 'VARIOS
                        'Me._cboAlmacen_2.SelectedValue(Index).Focus()
                        Me._cboAlmacen_2.Focus()
                End Select
            Case Else
                Select Case Index
                    Case 0 'JOYERIA
                        tecla = sender.KeyCode
                    Case 1 'RELOJERIA
                        tecla = sender.KeyCode
                    Case 2 'VARIOS
                        tecla = sender.KeyCode
                End Select
        End Select
    End Sub

    Private Sub _cboAlmacen_2_MouseUp(sender As Object, e As MouseEventArgs) Handles _cboAlmacen_2.MouseUp
        Dim Index As Integer
        '= _cboAlmacen_2.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._cboAlmacen_2.SelectedValue(Index).Text)
        Aux = Trim(Me._cboAlmacen_2.Text)
        'If Me._cboAlmacen_2.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._cboAlmacen_2.SelectedValue <> 0 Then
        '    '_cboAlmacen_2_Leave(_cboAlmacen_2.SelectedValue.Item(Index), New System.EventArgs())
        '   _cboAlmacen_2_Leave(_cboAlmacen_2.SelectedValue, New System.EventArgs())
        'End If

        'Me._cboAlmacen_2.SelectedValue(Index).Text = Aux
        Me._cboAlmacen_2.Text = Aux
    End Sub

    Private Sub _cboAlmacen_2_CursorChanged(sender As Object, e As EventArgs) Handles _cboAlmacen_2.CursorChanged
        Dim Index As Integer
        '= _cboAlmacen_2.SelectedValue(sender)
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub


        'lStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me._cboAlmacen_2.SelectedValue(Index).Text) & "%' Order by CodAlmacenOrigen "
        lStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me._cboAlmacen_2.Text) & "%' Order by CodAlmacenOrigen "

        Select Case Index
            Case 0 'JOYERIA
                'ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_2.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_2)
            Case 1 'RELOJERIA
                'ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_2.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_2)
            Case 2 'VARIOS
                'ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_2.SelectedValue(Index))
                ModDCombo.DCChange(lStrSql, tecla, _cboAlmacen_2)
        End Select


        'If Trim(Me._cboAlmacen_2.SelectedValue(Index).Text) = "" Then
        If Trim(Me._cboAlmacen_2.Text) = "" Then
            mintJOrigen = 0
            mintROrigen = 0
            mintVOrigen = 0
        End If
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub


    'Private Sub dbcProveedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcProveedor.CursorChanged
    '    Dim Index As Integer = dbcProveedor.SelectedValue(eventSender)
    '    Dim lStrSql As String


    '    lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.SelectedValue(Index).Text) & "%' Order by DescProvAcreed "
    '    ModDCombo.DCChange(lStrSql, tecla, dbcProveedor.SelectedValue(Index))


    '    If Trim(Me.dbcProveedor.SelectedValue(Index).Text) = "" Then
    '        mintJProv = 0
    '        mintRProv = 0
    '        mintVProv = 0
    '    End If
    'End Sub

    'Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
    '    Dim Index As Integer = dbcProveedor.SelectedValue(eventSender)
    '    Pon_Tool()
    '    gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' ORDER BY descProvAcreed "
    '    ModDCombo.DCGotFocus(gStrSql, dbcProveedor.SelectedValue(Index))
    'End Sub

    'Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcProveedor.KeyDown
    '    Dim Index As Integer = dbcProveedor.SelectedValue(eventSender)
    '    If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
    '        Select Case Index
    '            Case nJOYERIA
    '                Me.cboAlmacen.SelectedValue(nJOYERIA).Focus()
    '            Case nRELOJERIA
    '                Me.cboAlmacen.SelectedValue(Index).Focus()
    '            Case nVARIOS
    '                Me.cboAlmacen.SelectedValue(nVARIOS).Focus()
    '        End Select
    '        eventSender.KeyCode = 0
    '    End If
    '    tecla = eventArgs.KeyCode
    'End Sub

    'Private Sub dbcProveedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Leave
    '    Dim Index As Integer = dbcProveedor.SelectedValue(eventSender)
    '    Dim I As Integer
    '    Dim cDescripcion As String

    '    cDescripcion = Me.dbcProveedor.SelectedValue(Index).Text
    '    ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "%'")
    '    If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
    '        Exit Sub
    '    End If
    '    gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' and descProvAcreed = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
    '    mintJProv = 0
    '    mintRProv = 0
    '    mintVProv = 0
    '    Select Case Me.sstArticulo.SelectedIndex
    '        Case nJOYERIA
    '            ModDCombo.DCLostFocus(dbcProveedor.SelectedValue(Index), gStrSql, mintJProv)
    '        Case nRELOJERIA
    '            ModDCombo.DCLostFocus(dbcProveedor.SelectedValue(Index), gStrSql, mintRProv)
    '        Case nVARIOS
    '            ModDCombo.DCLostFocus(dbcProveedor.SelectedValue(Index), gStrSql, mintVProv)
    '    End Select
    '    If mintJProv = 0 And mintRProv = 0 And mintVProv = 0 Then
    '        mblnFueraChange = True

    '        Me.dbcProveedor.SelectedValue(Index).Text = cINDEFINIDO
    '        mblnFueraChange = False
    '    End If
    'End Sub

    'Private Sub dbcProveedor_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcProveedor.MouseUp
    '    Dim Index As Integer = dbcProveedor.SelectedValue(eventSender)
    '    Dim Aux As String

    '    Aux = Trim(Me.dbcProveedor.SelectedValue(Index).Text)
    '    If Me.dbcProveedor.SelectedValue(Index).SelectedItem <> 0 Then
    '        dbcProveedor_Leave(dbcProveedor.SelectedValue.Item(Index), New System.EventArgs())
    '    End If

    '    Me.dbcProveedor.SelectedValue(Index).Text = Aux
    'End Sub


    Private Sub _dbcProveedor_0_Enter(sender As Object, e As EventArgs) Handles _dbcProveedor_0.Enter
        Dim Index As Integer
        '= _dbcProveedor_0.SelectedValue(sender)
        Pon_Tool()
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' ORDER BY descProvAcreed "
        'ModDCombo.DCGotFocus(gStrSql, dbcProveedor.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _dbcProveedor_0)
    End Sub

    Private Sub _dbcProveedor_0_Leave(sender As Object, e As EventArgs) Handles _dbcProveedor_0.Leave
        Dim Index As Integer
        '= _dbcProveedor_0.SelectedValue(sender)
        Dim I As Integer
        Dim cDescripcion As String

        'cDescripcion = Me._dbcProveedor_0.SelectedValue(Index).Text
        cDescripcion = Me._dbcProveedor_0.Text
        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "%'")
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' and descProvAcreed = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
        mintJProv = 0
        mintRProv = 0
        mintVProv = 0
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                'ModDCombo.DCLostFocus(_dbcProveedor_0.SelectedValue(Index), gStrSql, mintJProv)
                ModDCombo.DCLostFocus(_dbcProveedor_0, gStrSql, mintJProv)
            Case nRELOJERIA
                'ModDCombo.DCLostFocus(_dbcProveedor_0.SelectedValue(Index), gStrSql, mintRProv)
                ModDCombo.DCLostFocus(_dbcProveedor_0, gStrSql, mintRProv)
            Case nVARIOS
                'ModDCombo.DCLostFocus(_dbcProveedor_0.SelectedValue(Index), gStrSql, mintVProv)
                ModDCombo.DCLostFocus(_dbcProveedor_0, gStrSql, mintVProv)
        End Select
        If mintJProv = 0 And mintRProv = 0 And mintVProv = 0 Then
            mblnFueraChange = True

            'Me._dbcProveedor_0.SelectedValue(Index).Text = cINDEFINIDO
            Me._dbcProveedor_0.Text = cINDEFINIDO
            mblnFueraChange = False
        End If
    End Sub

    Private Sub _dbcProveedor_0_KeyDown(sender As Object, e As KeyEventArgs) Handles _dbcProveedor_0.KeyDown
        Dim Index As Integer
        '= _dbcProveedor_0.SelectedValue(sender)
        If e.KeyCode = System.Windows.Forms.Keys.Escape Then
            Select Case Index
                Case nJOYERIA
                    Me._cboAlmacen_0.Focus()
                Case nRELOJERIA
                    Me._cboAlmacen_0.Focus()
                Case nVARIOS
                    Me._cboAlmacen_0.Focus()
            End Select
            sender.KeyCode = 0
        End If
        tecla = e.KeyCode
    End Sub

    Private Sub _dbcProveedor_0_MouseUp(sender As Object, e As MouseEventArgs) Handles _dbcProveedor_0.MouseUp
        Dim Index As Integer
        '= _dbcProveedor_0.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcProveedor_0.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcProveedor_0.Text)
        ''If Me._dbcProveedor_0.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._dbcProveedor_0.SelectedValue <> 0 Then
        '    '_dbcProveedor_0_Leave(_dbcProveedor_0.SelectedValue.Item(Index), New System.EventArgs())
        '    _dbcProveedor_0_Leave(_dbcProveedor_0.SelectedValue, New System.EventArgs())
        'End If

        'Me._dbcProveedor_0.SelectedValue(Index).Text = Aux
        Me._dbcProveedor_0.Text = Aux
    End Sub

    Private Sub _dbcProveedor_0_CursorChanged(sender As Object, e As EventArgs) Handles _dbcProveedor_0.CursorChanged
        Dim Index As Integer
        '= _dbcProveedor_0.SelectedValue(sender)
        Dim lStrSql As String


        'lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' and descProvAcreed LIKE '" & Trim(Me._dbcProveedor_0.SelectedValue(Index).Text) & "%' Order by DescProvAcreed "
        'ModDCombo.DCChange(lStrSql, tecla, _dbcProveedor_0.SelectedValue(Index))

        lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' and descProvAcreed LIKE '" & Trim(Me._dbcProveedor_0.Text) & "%' Order by DescProvAcreed "
        ModDCombo.DCChange(lStrSql, tecla, _dbcProveedor_0)


        'If Trim(Me._dbcProveedor_0.SelectedValue(Index).Text) = "" Then
        If Trim(Me._dbcProveedor_0.Text) = "" Then
            mintJProv = 0
            mintRProv = 0
            mintVProv = 0
        End If
    End Sub


    Private Sub _dbcProveedor_1_Enter(sender As Object, e As EventArgs) Handles _dbcProveedor_1.Enter
        Dim Index As Integer
        '= _dbcProveedor_1.SelectedValue(sender)
        Pon_Tool()
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' ORDER BY descProvAcreed "
        'ModDCombo.DCGotFocus(gStrSql, _dbcProveedor_1.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _dbcProveedor_1)
    End Sub

    Private Sub _dbcProveedor_1_Leave(sender As Object, e As EventArgs) Handles _dbcProveedor_1.Leave
        Dim Index As Integer
        '= _dbcProveedor_1.SelectedValue(sender)
        Dim I As Integer
        Dim cDescripcion As String

        'cDescripcion = Me._dbcProveedor_1.SelectedValue(Index).Text
        cDescripcion = Me._dbcProveedor_1.Text
        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "%'")
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' and descProvAcreed = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
        mintJProv = 0
        mintRProv = 0
        mintVProv = 0
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                'ModDCombo.DCLostFocus(_dbcProveedor_1.SelectedValue(Index), gStrSql, mintJProv)
                ModDCombo.DCLostFocus(_dbcProveedor_1, gStrSql, mintJProv)
            Case nRELOJERIA
                'ModDCombo.DCLostFocus(_dbcProveedor_1.SelectedValue(Index), gStrSql, mintRProv)
                ModDCombo.DCLostFocus(_dbcProveedor_1, gStrSql, mintRProv)
            Case nVARIOS
                'ModDCombo.DCLostFocus(_dbcProveedor_1.SelectedValue(Index), gStrSql, mintVProv)
                ModDCombo.DCLostFocus(_dbcProveedor_1, gStrSql, mintVProv)
        End Select
        If mintJProv = 0 And mintRProv = 0 And mintVProv = 0 Then
            mblnFueraChange = True

            'Me._dbcProveedor_1.SelectedValue(Index).Text = cINDEFINIDO
            Me._dbcProveedor_1.Text = cINDEFINIDO
            mblnFueraChange = False
        End If
    End Sub

    Private Sub _dbcProveedor_1_KeyDown(sender As Object, e As KeyEventArgs) Handles _dbcProveedor_1.KeyDown
        Dim Index As Integer
        '= _dbcProveedor_1.SelectedValue(sender)
        If e.KeyCode = System.Windows.Forms.Keys.Escape Then
            Select Case Index
                Case nJOYERIA
                    Me._cboAlmacen_1.Focus()
                Case nRELOJERIA
                    Me._cboAlmacen_1.Focus()
                Case nVARIOS
                    Me._cboAlmacen_1.Focus()
            End Select
            sender.KeyCode = 0
        End If
        tecla = e.KeyCode
    End Sub

    Private Sub _dbcProveedor_1_MouseUp(sender As Object, e As MouseEventArgs) Handles _dbcProveedor_1.MouseUp
        Dim Index As Integer
        '= _dbcProveedor_1.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcProveedor_1.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcProveedor_1.Text)
        ''If Me._dbcProveedor_1.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._dbcProveedor_1.SelectedValue <> 0 Then
        '    '_dbcProveedor_1_Leave(_dbcProveedor_1.SelectedValue.Item(Index), New System.EventArgs())
        '    _dbcProveedor_1_Leave(_dbcProveedor_1.SelectedValue, New System.EventArgs())
        'End If

        'Me._dbcProveedor_1.SelectedValue(Index).Text = Aux
        Me._dbcProveedor_1.Text = Aux
    End Sub

    Private Sub _dbcProveedor_1_CursorChanged(sender As Object, e As EventArgs) Handles _dbcProveedor_1.CursorChanged
        Dim Index As Integer
        '= _dbcProveedor_1.SelectedValue(sender)
        Dim lStrSql As String


        'lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' and descProvAcreed LIKE '" & Trim(Me._dbcProveedor_1.SelectedValue(Index).Text) & "%' Order by DescProvAcreed "
        'ModDCombo.DCChange(lStrSql, tecla, _dbcProveedor_1.SelectedValue(Index))

        lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' and descProvAcreed LIKE '" & Trim(Me._dbcProveedor_1.Text) & "%' Order by DescProvAcreed "
        ModDCombo.DCChange(lStrSql, tecla, _dbcProveedor_1)


        'If Trim(Me._dbcProveedor_1.SelectedValue(Index).Text) = "" Then
        If Trim(Me._dbcProveedor_1.Text) = "" Then
            mintJProv = 0
            mintRProv = 0
            mintVProv = 0
        End If
    End Sub

    Private Sub _dbcProveedor_2_Enter(sender As Object, e As EventArgs) Handles _dbcProveedor_2.Enter
        Dim Index As Integer
        '= _dbcProveedor_2.SelectedValue(sender)
        Pon_Tool()
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' ORDER BY descProvAcreed "
        'ModDCombo.DCGotFocus(gStrSql, _dbcProveedor_2.SelectedValue(Index))
        ModDCombo.DCGotFocus(gStrSql, _dbcProveedor_2)
    End Sub

    Private Sub _dbcProveedor_2_Leave(sender As Object, e As EventArgs) Handles _dbcProveedor_2.Leave
        Dim Index As Integer
        '= _dbcProveedor_2.SelectedValue(sender)
        Dim I As Integer
        Dim cDescripcion As String

        'cDescripcion = Me._dbcProveedor_2.SelectedValue(Index).Text
        cDescripcion = Me._dbcProveedor_2.Text
        ''" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "%'")
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' and descProvAcreed = '" & IIf(Trim(cDescripcion) = Trim(cINDEFINIDO), "'", Trim(cDescripcion) & "'")
        mintJProv = 0
        mintRProv = 0
        mintVProv = 0
        Select Case Me.sstArticulo.SelectedIndex
            Case nJOYERIA
                'ModDCombo.DCLostFocus(_dbcProveedor_2.SelectedValue(Index), gStrSql, mintJProv)
                ModDCombo.DCLostFocus(_dbcProveedor_2, gStrSql, mintJProv)
            Case nRELOJERIA
                'ModDCombo.DCLostFocus(_dbcProveedor_2.SelectedValue(Index), gStrSql, mintRProv)
                ModDCombo.DCLostFocus(_dbcProveedor_2, gStrSql, mintRProv)
            Case nVARIOS
                'ModDCombo.DCLostFocus(_dbcProveedor_2.SelectedValue(Index), gStrSql, mintVProv)
                ModDCombo.DCLostFocus(_dbcProveedor_2, gStrSql, mintVProv)
        End Select
        If mintJProv = 0 And mintRProv = 0 And mintVProv = 0 Then
            mblnFueraChange = True

            'Me._dbcProveedor_2.SelectedValue(Index).Text = cINDEFINIDO
            Me._dbcProveedor_2.Text = cINDEFINIDO
            mblnFueraChange = False
        End If
    End Sub

    Private Sub _dbcProveedor_2_KeyDown(sender As Object, e As KeyEventArgs) Handles _dbcProveedor_2.KeyDown
        Dim Index As Integer
        '= _dbcProveedor_2.SelectedValue(sender)
        If e.KeyCode = System.Windows.Forms.Keys.Escape Then
            Select Case Index
                Case nJOYERIA
                    Me._cboAlmacen_2.Focus()
                Case nRELOJERIA
                    Me._cboAlmacen_2.Focus()
                Case nVARIOS
                    Me._cboAlmacen_2.Focus()
            End Select
            sender.KeyCode = 0
        End If
        tecla = e.KeyCode
    End Sub

    Private Sub _dbcProveedor_2_MouseUp(sender As Object, e As MouseEventArgs) Handles _dbcProveedor_2.MouseUp
        Dim Index As Integer
        '= _dbcProveedor_2.SelectedValue(sender)
        Dim Aux As String

        'Aux = Trim(Me._dbcProveedor_2.SelectedValue(Index).Text)
        Aux = Trim(Me._dbcProveedor_2.Text)
        ''If Me._dbcProveedor_2.SelectedValue(Index).SelectedItem <> 0 Then
        'If Me._dbcProveedor_2.SelectedValue <> 0 Then
        '    '_dbcProveedor_2_Leave(_dbcProveedor_2.SelectedValue.Item(Index), New System.EventArgs())
        '    _dbcProveedor_2_Leave(_dbcProveedor_2.SelectedValue, New System.EventArgs())
        'End If

        'Me._dbcProveedor_2.SelectedValue(Index).Text = Aux
        Me._dbcProveedor_2.Text = Aux
    End Sub

    Private Sub _dbcProveedor_2_CursorChanged(sender As Object, e As EventArgs) Handles _dbcProveedor_2.CursorChanged
        Dim Index As Integer
        '= _dbcProveedor_2.SelectedValue(sender)
        Dim lStrSql As String


        'lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' and descProvAcreed LIKE '" & Trim(Me._dbcProveedor_2.SelectedValue(Index).Text) & "%' Order by DescProvAcreed "
        'ModDCombo.DCChange(lStrSql, tecla, _dbcProveedor_2.SelectedValue(Index))

        lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & Trim(C_TPROVEEDOR) & "' and descProvAcreed LIKE '" & Trim(Me._dbcProveedor_2.Text) & "%' Order by DescProvAcreed "
        ModDCombo.DCChange(lStrSql, tecla, _dbcProveedor_2)


        'If Trim(Me._dbcProveedor_2.SelectedValue(Index).Text) = "" Then
        If Trim(Me._dbcProveedor_2.Text) = "" Then
            mintJProv = 0
            mintRProv = 0
            mintVProv = 0
        End If
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Buscar()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtDescArticulo = New System.Windows.Forms.TextBox()
        Me.sstArticulo = New System.Windows.Forms.TabControl()
        Me._sstArticulo_TabPage0 = New System.Windows.Forms.TabPage()
        Me._fraContenedor_0 = New System.Windows.Forms.Panel()
        Me.fraDiamanteSuelto = New System.Windows.Forms.GroupBox()
        Me.txtMDSCertificado = New System.Windows.Forms.TextBox()
        Me.txtMDSPureza = New System.Windows.Forms.TextBox()
        Me.txtMDSColor = New System.Windows.Forms.TextBox()
        Me.txtMDSPeso = New System.Windows.Forms.TextBox()
        Me.lblEstatus = New System.Windows.Forms.Label()
        Me.lblMDSCertificado = New System.Windows.Forms.Label()
        Me.lblMDSPureza = New System.Windows.Forms.Label()
        Me.lblMDSColor = New System.Windows.Forms.Label()
        Me.lblMDSPeso = New System.Windows.Forms.Label()
        Me._txtAdicional_0 = New System.Windows.Forms.TextBox()
        Me._fraMoneda_5 = New System.Windows.Forms.GroupBox()
        Me._optMoneda_11 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_10 = New System.Windows.Forms.RadioButton()
        Me._fraMoneda_0 = New System.Windows.Forms.GroupBox()
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton()
        Me._fraImagen_0 = New System.Windows.Forms.GroupBox()
        Me.Image1 = New System.Windows.Forms.PictureBox()
        Me._Frame2_0 = New System.Windows.Forms.GroupBox()
        Me._Frame4_0 = New System.Windows.Forms.GroupBox()
        Me._cmdBuscarImagen_0 = New System.Windows.Forms.Button()
        Me._txtImagen_0 = New System.Windows.Forms.TextBox()
        Me._txtCodigodelProveedor_0 = New System.Windows.Forms.TextBox()
        Me._dbcProveedor_0 = New System.Windows.Forms.ComboBox()
        Me._cboUnidad_0 = New System.Windows.Forms.ComboBox()
        Me._cboAlmacen_0 = New System.Windows.Forms.ComboBox()
        Me._lblArticulo_36 = New System.Windows.Forms.Label()
        Me._lblArticulo_35 = New System.Windows.Forms.Label()
        Me._lblArticulo_11 = New System.Windows.Forms.Label()
        Me._lblArticulo_10 = New System.Windows.Forms.Label()
        Me._txtDescripcion_0 = New System.Windows.Forms.TextBox()
        Me._Frame1_0 = New System.Windows.Forms.GroupBox()
        Me._txtCostoReal_0 = New System.Windows.Forms.TextBox()
        Me._txtPrecioenDolares_0 = New System.Windows.Forms.TextBox()
        Me._txtCostoIndirecto_0 = New System.Windows.Forms.TextBox()
        Me._txtCostoAdicional_0 = New System.Windows.Forms.TextBox()
        Me._txtCostoFactura_0 = New System.Windows.Forms.TextBox()
        Me._lblMargen_0 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me._lblArticulo_34 = New System.Windows.Forms.Label()
        Me._lblArticulo_5 = New System.Windows.Forms.Label()
        Me._lblArticulo_8 = New System.Windows.Forms.Label()
        Me._lblArticulo_7 = New System.Windows.Forms.Label()
        Me._lblArticulo_6 = New System.Windows.Forms.Label()
        Me._dbcFamilia_0 = New System.Windows.Forms.ComboBox()
        Me._dbcLinea_0 = New System.Windows.Forms.ComboBox()
        Me.dbcSubLinea = New System.Windows.Forms.ComboBox()
        Me.dbcKilates = New System.Windows.Forms.ComboBox()
        Me._dbcMaterial_0 = New System.Windows.Forms.ComboBox()
        Me._lblArticulo_33 = New System.Windows.Forms.Label()
        Me._lblArticulo_9 = New System.Windows.Forms.Label()
        Me._lblArticulo_29 = New System.Windows.Forms.Label()
        Me._lblDescripcion_0 = New System.Windows.Forms.Label()
        Me._lblArticulo_26 = New System.Windows.Forms.Label()
        Me._lblArticulo_37 = New System.Windows.Forms.Label()
        Me._lblArticulo_4 = New System.Windows.Forms.Label()
        Me._lblArticulo_3 = New System.Windows.Forms.Label()
        Me._lblArticulo_2 = New System.Windows.Forms.Label()
        Me._lblArticulo_1 = New System.Windows.Forms.Label()
        Me._sstArticulo_TabPage1 = New System.Windows.Forms.TabPage()
        Me._fraContenedor_1 = New System.Windows.Forms.Panel()
        Me._txtAdicional_1 = New System.Windows.Forms.TextBox()
        Me._fraMoneda_3 = New System.Windows.Forms.GroupBox()
        Me._optMoneda_7 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_6 = New System.Windows.Forms.RadioButton()
        Me._txtDescripcion_1 = New System.Windows.Forms.TextBox()
        Me._fraArticulo_1 = New System.Windows.Forms.GroupBox()
        Me._optGenero_0 = New System.Windows.Forms.RadioButton()
        Me._optGenero_1 = New System.Windows.Forms.RadioButton()
        Me._optGenero_2 = New System.Windows.Forms.RadioButton()
        Me._fraArticulo_2 = New System.Windows.Forms.GroupBox()
        Me._optMovimiento_0 = New System.Windows.Forms.RadioButton()
        Me._optMovimiento_1 = New System.Windows.Forms.RadioButton()
        Me._optMovimiento_2 = New System.Windows.Forms.RadioButton()
        Me._fraImagen_1 = New System.Windows.Forms.GroupBox()
        Me.Image2 = New System.Windows.Forms.PictureBox()
        Me._fraMoneda_1 = New System.Windows.Forms.GroupBox()
        Me._optMoneda_3 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_2 = New System.Windows.Forms.RadioButton()
        Me._Frame1_2 = New System.Windows.Forms.GroupBox()
        Me._txtCostoReal_1 = New System.Windows.Forms.TextBox()
        Me._txtPrecioenDolares_1 = New System.Windows.Forms.TextBox()
        Me._txtCostoIndirecto_1 = New System.Windows.Forms.TextBox()
        Me._txtCostoAdicional_1 = New System.Windows.Forms.TextBox()
        Me._txtCostoFactura_1 = New System.Windows.Forms.TextBox()
        Me._lblMargen_1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me._lblArticulo_40 = New System.Windows.Forms.Label()
        Me._lblArticulo_41 = New System.Windows.Forms.Label()
        Me._lblArticulo_42 = New System.Windows.Forms.Label()
        Me._lblArticulo_43 = New System.Windows.Forms.Label()
        Me._lblArticulo_44 = New System.Windows.Forms.Label()
        Me._Frame2_2 = New System.Windows.Forms.GroupBox()
        Me._Frame4_1 = New System.Windows.Forms.GroupBox()
        Me._txtImagen_1 = New System.Windows.Forms.TextBox()
        Me._cmdBuscarImagen_1 = New System.Windows.Forms.Button()
        Me._txtCodigodelProveedor_1 = New System.Windows.Forms.TextBox()
        Me._dbcProveedor_1 = New System.Windows.Forms.ComboBox()
        Me._cboUnidad_1 = New System.Windows.Forms.ComboBox()
        Me._cboAlmacen_1 = New System.Windows.Forms.ComboBox()
        Me._lblArticulo_18 = New System.Windows.Forms.Label()
        Me._lblArticulo_19 = New System.Windows.Forms.Label()
        Me._lblArticulo_20 = New System.Windows.Forms.Label()
        Me._lblArticulo_21 = New System.Windows.Forms.Label()
        Me.chkCrono = New System.Windows.Forms.CheckBox()
        Me.dbcMarca = New System.Windows.Forms.ComboBox()
        Me.dbcModelo = New System.Windows.Forms.ComboBox()
        Me._dbcMaterial_1 = New System.Windows.Forms.ComboBox()
        Me._lblArticulo_45 = New System.Windows.Forms.Label()
        Me._lblArticulo_27 = New System.Windows.Forms.Label()
        Me._lblArticulo_13 = New System.Windows.Forms.Label()
        Me._lblArticulo_12 = New System.Windows.Forms.Label()
        Me._lblArticulo_14 = New System.Windows.Forms.Label()
        Me._lblArticulo_15 = New System.Windows.Forms.Label()
        Me._lblArticulo_16 = New System.Windows.Forms.Label()
        Me._lblArticulo_17 = New System.Windows.Forms.Label()
        Me._lblArticulo_38 = New System.Windows.Forms.Label()
        Me._lblDescripcion_1 = New System.Windows.Forms.Label()
        Me._sstArticulo_TabPage2 = New System.Windows.Forms.TabPage()
        Me._fraContenedor_2 = New System.Windows.Forms.Panel()
        Me._txtAdicional_2 = New System.Windows.Forms.TextBox()
        Me._fraMoneda_4 = New System.Windows.Forms.GroupBox()
        Me._optMoneda_9 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_8 = New System.Windows.Forms.RadioButton()
        Me._fraMoneda_2 = New System.Windows.Forms.GroupBox()
        Me._optMoneda_4 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_5 = New System.Windows.Forms.RadioButton()
        Me._Frame2_3 = New System.Windows.Forms.GroupBox()
        Me._Frame4_2 = New System.Windows.Forms.GroupBox()
        Me._txtImagen_2 = New System.Windows.Forms.TextBox()
        Me._cmdBuscarImagen_2 = New System.Windows.Forms.Button()
        Me._txtCodigodelProveedor_2 = New System.Windows.Forms.TextBox()
        Me._dbcProveedor_2 = New System.Windows.Forms.ComboBox()
        Me._cboUnidad_2 = New System.Windows.Forms.ComboBox()
        Me._cboAlmacen_2 = New System.Windows.Forms.ComboBox()
        Me._lblArticulo_50 = New System.Windows.Forms.Label()
        Me._lblArticulo_51 = New System.Windows.Forms.Label()
        Me._lblArticulo_52 = New System.Windows.Forms.Label()
        Me._lblArticulo_53 = New System.Windows.Forms.Label()
        Me._Frame1_3 = New System.Windows.Forms.GroupBox()
        Me._txtCostoFactura_2 = New System.Windows.Forms.TextBox()
        Me._txtCostoAdicional_2 = New System.Windows.Forms.TextBox()
        Me._txtCostoIndirecto_2 = New System.Windows.Forms.TextBox()
        Me._txtPrecioenDolares_2 = New System.Windows.Forms.TextBox()
        Me._txtCostoReal_2 = New System.Windows.Forms.TextBox()
        Me._lblMargen_2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me._lblArticulo_22 = New System.Windows.Forms.Label()
        Me._lblArticulo_23 = New System.Windows.Forms.Label()
        Me._lblArticulo_46 = New System.Windows.Forms.Label()
        Me._lblArticulo_47 = New System.Windows.Forms.Label()
        Me._lblArticulo_48 = New System.Windows.Forms.Label()
        Me._txtDescripcion_2 = New System.Windows.Forms.TextBox()
        Me._fraImagen_2 = New System.Windows.Forms.GroupBox()
        Me.Image3 = New System.Windows.Forms.PictureBox()
        Me._dbcFamilia_1 = New System.Windows.Forms.ComboBox()
        Me._dbcLinea_1 = New System.Windows.Forms.ComboBox()
        Me._dbcMaterial_2 = New System.Windows.Forms.ComboBox()
        Me._lblArticulo_54 = New System.Windows.Forms.Label()
        Me._lblArticulo_49 = New System.Windows.Forms.Label()
        Me._lblArticulo_28 = New System.Windows.Forms.Label()
        Me._lblArticulo_24 = New System.Windows.Forms.Label()
        Me._lblArticulo_25 = New System.Windows.Forms.Label()
        Me._lblArticulo_30 = New System.Windows.Forms.Label()
        Me._lblArticulo_39 = New System.Windows.Forms.Label()
        Me._lblDescripcion_2 = New System.Windows.Forms.Label()
        Me.chkCodigoAnterior = New System.Windows.Forms.CheckBox()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.txtCodArtAnterior = New System.Windows.Forms.TextBox()
        Me.dbcOrigen = New System.Windows.Forms.ComboBox()
        Me._lblArticulo_32 = New System.Windows.Forms.Label()
        Me._lblArticulo_31 = New System.Windows.Forms.Label()
        Me.txtCodArticulo = New System.Windows.Forms.TextBox()
        Me._lblArticulo_0 = New System.Windows.Forms.Label()
        Me.Frame1 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.Frame2 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.Frame4 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.cboAlmacen = New System.Windows.Forms.ComboBox()
        Me.cboUnidad = New System.Windows.Forms.ComboBox()
        Me.cmdBuscarImagen = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.dbcFamilia = New System.Windows.Forms.ComboBox()
        Me.dbcLinea = New System.Windows.Forms.ComboBox()
        Me.dbcMaterial = New System.Windows.Forms.ComboBox()
        Me.dbcProveedor = New System.Windows.Forms.ComboBox()
        Me.fraArticulo = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.fraContenedor = New Microsoft.VisualBasic.Compatibility.VB6.PanelArray(Me.components)
        Me.fraImagen = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.fraMoneda = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblArticulo = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblDescripcion = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblMargen = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optGenero = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.optMovimiento = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.txtAdicional = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtCodigodelProveedor = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtCostoAdicional = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtCostoFactura = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtCostoIndirecto = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtCostoReal = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtDescripcion = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtImagen = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.txtPrecioenDolares = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.sstArticulo.SuspendLayout()
        Me._sstArticulo_TabPage0.SuspendLayout()
        Me._fraContenedor_0.SuspendLayout()
        Me.fraDiamanteSuelto.SuspendLayout()
        Me._fraMoneda_5.SuspendLayout()
        Me._fraMoneda_0.SuspendLayout()
        Me._fraImagen_0.SuspendLayout()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._Frame2_0.SuspendLayout()
        Me._Frame4_0.SuspendLayout()
        Me._Frame1_0.SuspendLayout()
        Me._sstArticulo_TabPage1.SuspendLayout()
        Me._fraContenedor_1.SuspendLayout()
        Me._fraMoneda_3.SuspendLayout()
        Me._fraArticulo_1.SuspendLayout()
        Me._fraArticulo_2.SuspendLayout()
        Me._fraImagen_1.SuspendLayout()
        CType(Me.Image2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._fraMoneda_1.SuspendLayout()
        Me._Frame1_2.SuspendLayout()
        Me._Frame2_2.SuspendLayout()
        Me._Frame4_1.SuspendLayout()
        Me._sstArticulo_TabPage2.SuspendLayout()
        Me._fraContenedor_2.SuspendLayout()
        Me._fraMoneda_4.SuspendLayout()
        Me._fraMoneda_2.SuspendLayout()
        Me._Frame2_3.SuspendLayout()
        Me._Frame4_2.SuspendLayout()
        Me._Frame1_3.SuspendLayout()
        Me._fraImagen_2.SuspendLayout()
        CType(Me.Image3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame3.SuspendLayout()
        CType(Me.Frame1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Frame2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Frame4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdBuscarImagen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraArticulo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraContenedor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraImagen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblArticulo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDescripcion, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblMargen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optGenero, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMovimiento, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAdicional, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCodigodelProveedor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCostoAdicional, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCostoFactura, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCostoIndirecto, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCostoReal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDescripcion, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtImagen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPrecioenDolares, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtDescArticulo
        '
        Me.txtDescArticulo.AcceptsReturn = True
        Me.txtDescArticulo.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescArticulo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescArticulo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescArticulo.Location = New System.Drawing.Point(168, 8)
        Me.txtDescArticulo.MaxLength = 0
        Me.txtDescArticulo.Name = "txtDescArticulo"
        Me.txtDescArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescArticulo.Size = New System.Drawing.Size(328, 20)
        Me.txtDescArticulo.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtDescArticulo, "Descripción del artículo")
        '
        'sstArticulo
        '
        Me.sstArticulo.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.sstArticulo.Controls.Add(Me._sstArticulo_TabPage0)
        Me.sstArticulo.Controls.Add(Me._sstArticulo_TabPage1)
        Me.sstArticulo.Controls.Add(Me._sstArticulo_TabPage2)
        Me.sstArticulo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.sstArticulo.ItemSize = New System.Drawing.Size(42, 18)
        Me.sstArticulo.Location = New System.Drawing.Point(12, 64)
        Me.sstArticulo.Name = "sstArticulo"
        Me.sstArticulo.SelectedIndex = 0
        Me.sstArticulo.Size = New System.Drawing.Size(708, 567)
        Me.sstArticulo.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.sstArticulo, "Grupo de Artículo al que pertenece")
        '
        '_sstArticulo_TabPage0
        '
        Me._sstArticulo_TabPage0.Controls.Add(Me._fraContenedor_0)
        Me._sstArticulo_TabPage0.Location = New System.Drawing.Point(4, 22)
        Me._sstArticulo_TabPage0.Name = "_sstArticulo_TabPage0"
        Me._sstArticulo_TabPage0.Size = New System.Drawing.Size(700, 541)
        Me._sstArticulo_TabPage0.TabIndex = 0
        Me._sstArticulo_TabPage0.Text = "Joyería"
        '
        '_fraContenedor_0
        '
        Me._fraContenedor_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraContenedor_0.Controls.Add(Me.fraDiamanteSuelto)
        Me._fraContenedor_0.Controls.Add(Me._txtAdicional_0)
        Me._fraContenedor_0.Controls.Add(Me._fraMoneda_5)
        Me._fraContenedor_0.Controls.Add(Me._fraMoneda_0)
        Me._fraContenedor_0.Controls.Add(Me._fraImagen_0)
        Me._fraContenedor_0.Controls.Add(Me._Frame2_0)
        Me._fraContenedor_0.Controls.Add(Me._txtDescripcion_0)
        Me._fraContenedor_0.Controls.Add(Me._Frame1_0)
        Me._fraContenedor_0.Controls.Add(Me._dbcFamilia_0)
        Me._fraContenedor_0.Controls.Add(Me._dbcLinea_0)
        Me._fraContenedor_0.Controls.Add(Me.dbcSubLinea)
        Me._fraContenedor_0.Controls.Add(Me.dbcKilates)
        Me._fraContenedor_0.Controls.Add(Me._dbcMaterial_0)
        Me._fraContenedor_0.Controls.Add(Me._lblArticulo_33)
        Me._fraContenedor_0.Controls.Add(Me._lblArticulo_9)
        Me._fraContenedor_0.Controls.Add(Me._lblArticulo_29)
        Me._fraContenedor_0.Controls.Add(Me._lblDescripcion_0)
        Me._fraContenedor_0.Controls.Add(Me._lblArticulo_26)
        Me._fraContenedor_0.Controls.Add(Me._lblArticulo_37)
        Me._fraContenedor_0.Controls.Add(Me._lblArticulo_4)
        Me._fraContenedor_0.Controls.Add(Me._lblArticulo_3)
        Me._fraContenedor_0.Controls.Add(Me._lblArticulo_2)
        Me._fraContenedor_0.Controls.Add(Me._lblArticulo_1)
        Me._fraContenedor_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._fraContenedor_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraContenedor_0.Location = New System.Drawing.Point(8, 24)
        Me._fraContenedor_0.Name = "_fraContenedor_0"
        Me._fraContenedor_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraContenedor_0.Size = New System.Drawing.Size(680, 504)
        Me._fraContenedor_0.TabIndex = 10
        '
        'fraDiamanteSuelto
        '
        Me.fraDiamanteSuelto.BackColor = System.Drawing.SystemColors.Control
        Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSCertificado)
        Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSPureza)
        Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSColor)
        Me.fraDiamanteSuelto.Controls.Add(Me.txtMDSPeso)
        Me.fraDiamanteSuelto.Controls.Add(Me.lblEstatus)
        Me.fraDiamanteSuelto.Controls.Add(Me.lblMDSCertificado)
        Me.fraDiamanteSuelto.Controls.Add(Me.lblMDSPureza)
        Me.fraDiamanteSuelto.Controls.Add(Me.lblMDSColor)
        Me.fraDiamanteSuelto.Controls.Add(Me.lblMDSPeso)
        Me.fraDiamanteSuelto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraDiamanteSuelto.Location = New System.Drawing.Point(236, 102)
        Me.fraDiamanteSuelto.Name = "fraDiamanteSuelto"
        Me.fraDiamanteSuelto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraDiamanteSuelto.Size = New System.Drawing.Size(268, 92)
        Me.fraDiamanteSuelto.TabIndex = 23
        Me.fraDiamanteSuelto.TabStop = False
        Me.fraDiamanteSuelto.Text = " "
        '
        'txtMDSCertificado
        '
        Me.txtMDSCertificado.AcceptsReturn = True
        Me.txtMDSCertificado.BackColor = System.Drawing.SystemColors.Window
        Me.txtMDSCertificado.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMDSCertificado.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMDSCertificado.Location = New System.Drawing.Point(125, 40)
        Me.txtMDSCertificado.MaxLength = 20
        Me.txtMDSCertificado.Name = "txtMDSCertificado"
        Me.txtMDSCertificado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMDSCertificado.Size = New System.Drawing.Size(136, 20)
        Me.txtMDSCertificado.TabIndex = 31
        '
        'txtMDSPureza
        '
        Me.txtMDSPureza.AcceptsReturn = True
        Me.txtMDSPureza.BackColor = System.Drawing.SystemColors.Window
        Me.txtMDSPureza.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMDSPureza.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMDSPureza.Location = New System.Drawing.Point(71, 66)
        Me.txtMDSPureza.MaxLength = 4
        Me.txtMDSPureza.Name = "txtMDSPureza"
        Me.txtMDSPureza.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMDSPureza.Size = New System.Drawing.Size(50, 20)
        Me.txtMDSPureza.TabIndex = 29
        '
        'txtMDSColor
        '
        Me.txtMDSColor.AcceptsReturn = True
        Me.txtMDSColor.BackColor = System.Drawing.SystemColors.Window
        Me.txtMDSColor.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMDSColor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMDSColor.Location = New System.Drawing.Point(71, 40)
        Me.txtMDSColor.MaxLength = 1
        Me.txtMDSColor.Name = "txtMDSColor"
        Me.txtMDSColor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMDSColor.Size = New System.Drawing.Size(50, 20)
        Me.txtMDSColor.TabIndex = 27
        '
        'txtMDSPeso
        '
        Me.txtMDSPeso.AcceptsReturn = True
        Me.txtMDSPeso.BackColor = System.Drawing.SystemColors.Window
        Me.txtMDSPeso.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMDSPeso.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMDSPeso.Location = New System.Drawing.Point(71, 14)
        Me.txtMDSPeso.MaxLength = 6
        Me.txtMDSPeso.Name = "txtMDSPeso"
        Me.txtMDSPeso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMDSPeso.Size = New System.Drawing.Size(50, 20)
        Me.txtMDSPeso.TabIndex = 25
        Me.txtMDSPeso.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblEstatus
        '
        Me.lblEstatus.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblEstatus.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEstatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEstatus.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblEstatus.Location = New System.Drawing.Point(125, 68)
        Me.lblEstatus.Name = "lblEstatus"
        Me.lblEstatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEstatus.Size = New System.Drawing.Size(136, 19)
        Me.lblEstatus.TabIndex = 32
        Me.lblEstatus.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblMDSCertificado
        '
        Me.lblMDSCertificado.BackColor = System.Drawing.SystemColors.Control
        Me.lblMDSCertificado.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMDSCertificado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMDSCertificado.Location = New System.Drawing.Point(127, 17)
        Me.lblMDSCertificado.Name = "lblMDSCertificado"
        Me.lblMDSCertificado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMDSCertificado.Size = New System.Drawing.Size(63, 18)
        Me.lblMDSCertificado.TabIndex = 30
        Me.lblMDSCertificado.Text = "Certificado"
        '
        'lblMDSPureza
        '
        Me.lblMDSPureza.BackColor = System.Drawing.SystemColors.Control
        Me.lblMDSPureza.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMDSPureza.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMDSPureza.Location = New System.Drawing.Point(9, 71)
        Me.lblMDSPureza.Name = "lblMDSPureza"
        Me.lblMDSPureza.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMDSPureza.Size = New System.Drawing.Size(63, 16)
        Me.lblMDSPureza.TabIndex = 28
        Me.lblMDSPureza.Text = "Pureza - Q"
        '
        'lblMDSColor
        '
        Me.lblMDSColor.BackColor = System.Drawing.SystemColors.Control
        Me.lblMDSColor.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMDSColor.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMDSColor.Location = New System.Drawing.Point(9, 44)
        Me.lblMDSColor.Name = "lblMDSColor"
        Me.lblMDSColor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMDSColor.Size = New System.Drawing.Size(63, 18)
        Me.lblMDSColor.TabIndex = 26
        Me.lblMDSColor.Text = "Color"
        '
        'lblMDSPeso
        '
        Me.lblMDSPeso.BackColor = System.Drawing.SystemColors.Control
        Me.lblMDSPeso.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMDSPeso.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMDSPeso.Location = New System.Drawing.Point(9, 18)
        Me.lblMDSPeso.Name = "lblMDSPeso"
        Me.lblMDSPeso.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMDSPeso.Size = New System.Drawing.Size(63, 18)
        Me.lblMDSPeso.TabIndex = 24
        Me.lblMDSPeso.Text = "Peso - CT"
        '
        '_txtAdicional_0
        '
        Me._txtAdicional_0.AcceptsReturn = True
        Me._txtAdicional_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
        Me._txtAdicional_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtAdicional_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtAdicional_0.Location = New System.Drawing.Point(89, 168)
        Me._txtAdicional_0.MaxLength = 15
        Me._txtAdicional_0.Name = "_txtAdicional_0"
        Me._txtAdicional_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtAdicional_0.Size = New System.Drawing.Size(120, 20)
        Me._txtAdicional_0.TabIndex = 22
        '
        '_fraMoneda_5
        '
        Me._fraMoneda_5.BackColor = System.Drawing.SystemColors.Control
        Me._fraMoneda_5.Controls.Add(Me._optMoneda_11)
        Me._fraMoneda_5.Controls.Add(Me._optMoneda_10)
        Me._fraMoneda_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraMoneda_5.Location = New System.Drawing.Point(89, 256)
        Me._fraMoneda_5.Name = "_fraMoneda_5"
        Me._fraMoneda_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraMoneda_5.Size = New System.Drawing.Size(218, 33)
        Me._fraMoneda_5.TabIndex = 37
        Me._fraMoneda_5.TabStop = False
        '
        '_optMoneda_11
        '
        Me._optMoneda_11.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_11.Location = New System.Drawing.Point(129, 11)
        Me._optMoneda_11.Name = "_optMoneda_11"
        Me._optMoneda_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_11.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_11.TabIndex = 39
        Me._optMoneda_11.TabStop = True
        Me._optMoneda_11.Tag = "0"
        Me._optMoneda_11.Text = "Pesos"
        Me._optMoneda_11.UseVisualStyleBackColor = False
        '
        '_optMoneda_10
        '
        Me._optMoneda_10.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_10.Checked = True
        Me._optMoneda_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_10.Location = New System.Drawing.Point(36, 11)
        Me._optMoneda_10.Name = "_optMoneda_10"
        Me._optMoneda_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_10.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_10.TabIndex = 38
        Me._optMoneda_10.TabStop = True
        Me._optMoneda_10.Tag = "1"
        Me._optMoneda_10.Text = "Dólares"
        Me._optMoneda_10.UseVisualStyleBackColor = False
        '
        '_fraMoneda_0
        '
        Me._fraMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraMoneda_0.Controls.Add(Me._optMoneda_1)
        Me._fraMoneda_0.Controls.Add(Me._optMoneda_0)
        Me._fraMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraMoneda_0.Location = New System.Drawing.Point(416, 256)
        Me._fraMoneda_0.Name = "_fraMoneda_0"
        Me._fraMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraMoneda_0.Size = New System.Drawing.Size(209, 33)
        Me._fraMoneda_0.TabIndex = 41
        Me._fraMoneda_0.TabStop = False
        '
        '_optMoneda_1
        '
        Me._optMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_1.Location = New System.Drawing.Point(134, 11)
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_1.TabIndex = 43
        Me._optMoneda_1.TabStop = True
        Me._optMoneda_1.Text = "Pesos"
        Me._optMoneda_1.UseVisualStyleBackColor = False
        '
        '_optMoneda_0
        '
        Me._optMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_0.Checked = True
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_0.Location = New System.Drawing.Point(34, 11)
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_0.TabIndex = 42
        Me._optMoneda_0.TabStop = True
        Me._optMoneda_0.Text = "Dólares"
        Me._optMoneda_0.UseVisualStyleBackColor = False
        '
        '_fraImagen_0
        '
        Me._fraImagen_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraImagen_0.Controls.Add(Me.Image1)
        Me._fraImagen_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraImagen_0.Location = New System.Drawing.Point(510, 8)
        Me._fraImagen_0.Name = "_fraImagen_0"
        Me._fraImagen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraImagen_0.Size = New System.Drawing.Size(178, 186)
        Me._fraImagen_0.TabIndex = 69
        Me._fraImagen_0.TabStop = False
        Me._fraImagen_0.Text = "Imagen del Artículo"
        '
        'Image1
        '
        Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image1.Image = Global.CorporativoV1.My.Resources.Resources.JMR2
        Me.Image1.Location = New System.Drawing.Point(7, 21)
        Me.Image1.Name = "Image1"
        Me.Image1.Size = New System.Drawing.Size(163, 157)
        Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Image1.TabIndex = 0
        Me.Image1.TabStop = False
        '
        '_Frame2_0
        '
        Me._Frame2_0.BackColor = System.Drawing.SystemColors.Control
        Me._Frame2_0.Controls.Add(Me._Frame4_0)
        Me._Frame2_0.Controls.Add(Me._txtCodigodelProveedor_0)
        Me._Frame2_0.Controls.Add(Me._dbcProveedor_0)
        Me._Frame2_0.Controls.Add(Me._cboUnidad_0)
        Me._Frame2_0.Controls.Add(Me._cboAlmacen_0)
        Me._Frame2_0.Controls.Add(Me._lblArticulo_36)
        Me._Frame2_0.Controls.Add(Me._lblArticulo_35)
        Me._Frame2_0.Controls.Add(Me._lblArticulo_11)
        Me._Frame2_0.Controls.Add(Me._lblArticulo_10)
        Me._Frame2_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame2_0.Location = New System.Drawing.Point(312, 294)
        Me._Frame2_0.Name = "_Frame2_0"
        Me._Frame2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame2_0.Size = New System.Drawing.Size(313, 185)
        Me._Frame2_0.TabIndex = 57
        Me._Frame2_0.TabStop = False
        '
        '_Frame4_0
        '
        Me._Frame4_0.BackColor = System.Drawing.SystemColors.Control
        Me._Frame4_0.Controls.Add(Me._cmdBuscarImagen_0)
        Me._Frame4_0.Controls.Add(Me._txtImagen_0)
        Me._Frame4_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._Frame4_0.Location = New System.Drawing.Point(12, 132)
        Me._Frame4_0.Name = "_Frame4_0"
        Me._Frame4_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame4_0.Size = New System.Drawing.Size(290, 44)
        Me._Frame4_0.TabIndex = 66
        Me._Frame4_0.TabStop = False
        Me._Frame4_0.Text = "Imagen"
        '
        '_cmdBuscarImagen_0
        '
        Me._cmdBuscarImagen_0.BackColor = System.Drawing.SystemColors.Control
        Me._cmdBuscarImagen_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdBuscarImagen_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._cmdBuscarImagen_0.Location = New System.Drawing.Point(260, 15)
        Me._cmdBuscarImagen_0.Name = "_cmdBuscarImagen_0"
        Me._cmdBuscarImagen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdBuscarImagen_0.Size = New System.Drawing.Size(22, 21)
        Me._cmdBuscarImagen_0.TabIndex = 68
        Me._cmdBuscarImagen_0.Text = "..."
        Me._cmdBuscarImagen_0.UseVisualStyleBackColor = False
        '
        '_txtImagen_0
        '
        Me._txtImagen_0.AcceptsReturn = True
        Me._txtImagen_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtImagen_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtImagen_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtImagen_0.Location = New System.Drawing.Point(9, 15)
        Me._txtImagen_0.MaxLength = 0
        Me._txtImagen_0.Name = "_txtImagen_0"
        Me._txtImagen_0.ReadOnly = True
        Me._txtImagen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtImagen_0.Size = New System.Drawing.Size(245, 20)
        Me._txtImagen_0.TabIndex = 67
        '
        '_txtCodigodelProveedor_0
        '
        Me._txtCodigodelProveedor_0.AcceptsReturn = True
        Me._txtCodigodelProveedor_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
        Me._txtCodigodelProveedor_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCodigodelProveedor_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtCodigodelProveedor_0.Location = New System.Drawing.Point(171, 102)
        Me._txtCodigodelProveedor_0.MaxLength = 20
        Me._txtCodigodelProveedor_0.Name = "_txtCodigodelProveedor_0"
        Me._txtCodigodelProveedor_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCodigodelProveedor_0.Size = New System.Drawing.Size(129, 20)
        Me._txtCodigodelProveedor_0.TabIndex = 65
        Me.ToolTip1.SetToolTip(Me._txtCodigodelProveedor_0, "Código que usa el Proveedor para el Artículo")
        '
        '_dbcProveedor_0
        '
        Me._dbcProveedor_0.Location = New System.Drawing.Point(100, 74)
        Me._dbcProveedor_0.Name = "_dbcProveedor_0"
        Me._dbcProveedor_0.Size = New System.Drawing.Size(201, 21)
        Me._dbcProveedor_0.TabIndex = 63
        '
        '_cboUnidad_0
        '
        Me._cboUnidad_0.Location = New System.Drawing.Point(100, 17)
        Me._cboUnidad_0.Name = "_cboUnidad_0"
        Me._cboUnidad_0.Size = New System.Drawing.Size(78, 21)
        Me._cboUnidad_0.TabIndex = 59
        '
        '_cboAlmacen_0
        '
        Me._cboAlmacen_0.Location = New System.Drawing.Point(100, 46)
        Me._cboAlmacen_0.Name = "_cboAlmacen_0"
        Me._cboAlmacen_0.Size = New System.Drawing.Size(201, 21)
        Me._cboAlmacen_0.TabIndex = 61
        '
        '_lblArticulo_36
        '
        Me._lblArticulo_36.AutoSize = True
        Me._lblArticulo_36.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_36.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_36.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_36.Location = New System.Drawing.Point(12, 50)
        Me._lblArticulo_36.Name = "_lblArticulo_36"
        Me._lblArticulo_36.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_36.Size = New System.Drawing.Size(84, 13)
        Me._lblArticulo_36.TabIndex = 60
        Me._lblArticulo_36.Text = "Almacén/Origen"
        '
        '_lblArticulo_35
        '
        Me._lblArticulo_35.AutoSize = True
        Me._lblArticulo_35.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_35.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_35.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_35.Location = New System.Drawing.Point(12, 21)
        Me._lblArticulo_35.Name = "_lblArticulo_35"
        Me._lblArticulo_35.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_35.Size = New System.Drawing.Size(41, 13)
        Me._lblArticulo_35.TabIndex = 58
        Me._lblArticulo_35.Text = "Unidad"
        '
        '_lblArticulo_11
        '
        Me._lblArticulo_11.AutoSize = True
        Me._lblArticulo_11.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_11.Location = New System.Drawing.Point(19, 106)
        Me._lblArticulo_11.Name = "_lblArticulo_11"
        Me._lblArticulo_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_11.Size = New System.Drawing.Size(156, 13)
        Me._lblArticulo_11.TabIndex = 64
        Me._lblArticulo_11.Text = "Código artículo del proveedor : "
        '
        '_lblArticulo_10
        '
        Me._lblArticulo_10.AutoSize = True
        Me._lblArticulo_10.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_10.Location = New System.Drawing.Point(12, 79)
        Me._lblArticulo_10.Name = "_lblArticulo_10"
        Me._lblArticulo_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_10.Size = New System.Drawing.Size(56, 13)
        Me._lblArticulo_10.TabIndex = 62
        Me._lblArticulo_10.Text = "Proveedor"
        '
        '_txtDescripcion_0
        '
        Me._txtDescripcion_0.AcceptsReturn = True
        Me._txtDescripcion_0.BackColor = System.Drawing.SystemColors.Info
        Me._txtDescripcion_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtDescripcion_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
        Me._txtDescripcion_0.Location = New System.Drawing.Point(89, 198)
        Me._txtDescripcion_0.MaxLength = 150
        Me._txtDescripcion_0.Name = "_txtDescripcion_0"
        Me._txtDescripcion_0.ReadOnly = True
        Me._txtDescripcion_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtDescripcion_0.Size = New System.Drawing.Size(537, 20)
        Me._txtDescripcion_0.TabIndex = 34
        Me.ToolTip1.SetToolTip(Me._txtDescripcion_0, "Descripción del Artículo")
        '
        '_Frame1_0
        '
        Me._Frame1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Frame1_0.Controls.Add(Me._txtCostoReal_0)
        Me._Frame1_0.Controls.Add(Me._txtPrecioenDolares_0)
        Me._Frame1_0.Controls.Add(Me._txtCostoIndirecto_0)
        Me._Frame1_0.Controls.Add(Me._txtCostoAdicional_0)
        Me._Frame1_0.Controls.Add(Me._txtCostoFactura_0)
        Me._Frame1_0.Controls.Add(Me._lblMargen_0)
        Me._Frame1_0.Controls.Add(Me.Label1)
        Me._Frame1_0.Controls.Add(Me._lblArticulo_34)
        Me._Frame1_0.Controls.Add(Me._lblArticulo_5)
        Me._Frame1_0.Controls.Add(Me._lblArticulo_8)
        Me._Frame1_0.Controls.Add(Me._lblArticulo_7)
        Me._Frame1_0.Controls.Add(Me._lblArticulo_6)
        Me._Frame1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame1_0.Location = New System.Drawing.Point(5, 294)
        Me._Frame1_0.Name = "_Frame1_0"
        Me._Frame1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_0.Size = New System.Drawing.Size(300, 185)
        Me._Frame1_0.TabIndex = 44
        Me._Frame1_0.TabStop = False
        '
        '_txtCostoReal_0
        '
        Me._txtCostoReal_0.AcceptsReturn = True
        Me._txtCostoReal_0.BackColor = System.Drawing.SystemColors.Info
        Me._txtCostoReal_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCostoReal_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
        Me._txtCostoReal_0.Location = New System.Drawing.Point(92, 152)
        Me._txtCostoReal_0.MaxLength = 0
        Me._txtCostoReal_0.Name = "_txtCostoReal_0"
        Me._txtCostoReal_0.ReadOnly = True
        Me._txtCostoReal_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCostoReal_0.Size = New System.Drawing.Size(113, 20)
        Me._txtCostoReal_0.TabIndex = 54
        Me._txtCostoReal_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtCostoReal_0, "Costo Real del artículo")
        '
        '_txtPrecioenDolares_0
        '
        Me._txtPrecioenDolares_0.AcceptsReturn = True
        Me._txtPrecioenDolares_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me._txtPrecioenDolares_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtPrecioenDolares_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtPrecioenDolares_0.Location = New System.Drawing.Point(92, 24)
        Me._txtPrecioenDolares_0.MaxLength = 0
        Me._txtPrecioenDolares_0.Name = "_txtPrecioenDolares_0"
        Me._txtPrecioenDolares_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtPrecioenDolares_0.Size = New System.Drawing.Size(113, 20)
        Me._txtPrecioenDolares_0.TabIndex = 46
        Me._txtPrecioenDolares_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtPrecioenDolares_0, "Precio al Público en Dólares")
        '
        '_txtCostoIndirecto_0
        '
        Me._txtCostoIndirecto_0.AcceptsReturn = True
        Me._txtCostoIndirecto_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtCostoIndirecto_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCostoIndirecto_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._txtCostoIndirecto_0.Location = New System.Drawing.Point(92, 120)
        Me._txtCostoIndirecto_0.MaxLength = 0
        Me._txtCostoIndirecto_0.Name = "_txtCostoIndirecto_0"
        Me._txtCostoIndirecto_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCostoIndirecto_0.Size = New System.Drawing.Size(113, 20)
        Me._txtCostoIndirecto_0.TabIndex = 52
        Me._txtCostoIndirecto_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtCostoIndirecto_0, "Gastos Indirectos en Dólares")
        '
        '_txtCostoAdicional_0
        '
        Me._txtCostoAdicional_0.AcceptsReturn = True
        Me._txtCostoAdicional_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtCostoAdicional_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCostoAdicional_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._txtCostoAdicional_0.Location = New System.Drawing.Point(92, 88)
        Me._txtCostoAdicional_0.MaxLength = 0
        Me._txtCostoAdicional_0.Name = "_txtCostoAdicional_0"
        Me._txtCostoAdicional_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCostoAdicional_0.Size = New System.Drawing.Size(113, 20)
        Me._txtCostoAdicional_0.TabIndex = 50
        Me._txtCostoAdicional_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtCostoAdicional_0, "Costo en Dólares")
        '
        '_txtCostoFactura_0
        '
        Me._txtCostoFactura_0.AcceptsReturn = True
        Me._txtCostoFactura_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtCostoFactura_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCostoFactura_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._txtCostoFactura_0.Location = New System.Drawing.Point(92, 56)
        Me._txtCostoFactura_0.MaxLength = 0
        Me._txtCostoFactura_0.Name = "_txtCostoFactura_0"
        Me._txtCostoFactura_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCostoFactura_0.Size = New System.Drawing.Size(113, 20)
        Me._txtCostoFactura_0.TabIndex = 48
        Me._txtCostoFactura_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtCostoFactura_0, "Costo en Pesos")
        '
        '_lblMargen_0
        '
        Me._lblMargen_0.BackColor = System.Drawing.SystemColors.Window
        Me._lblMargen_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._lblMargen_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblMargen_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblMargen_0.Location = New System.Drawing.Point(244, 152)
        Me._lblMargen_0.Name = "_lblMargen_0"
        Me._lblMargen_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblMargen_0.Size = New System.Drawing.Size(49, 21)
        Me._lblMargen_0.TabIndex = 56
        Me._lblMargen_0.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(232, 120)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(61, 29)
        Me.Label1.TabIndex = 55
        Me.Label1.Text = "% Margen de Venta "
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblArticulo_34
        '
        Me._lblArticulo_34.AutoSize = True
        Me._lblArticulo_34.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_34.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_34.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_34.Location = New System.Drawing.Point(12, 156)
        Me._lblArticulo_34.Name = "_lblArticulo_34"
        Me._lblArticulo_34.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_34.Size = New System.Drawing.Size(59, 13)
        Me._lblArticulo_34.TabIndex = 53
        Me._lblArticulo_34.Text = "Costo Real"
        '
        '_lblArticulo_5
        '
        Me._lblArticulo_5.AutoSize = True
        Me._lblArticulo_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_5.Location = New System.Drawing.Point(12, 28)
        Me._lblArticulo_5.Name = "_lblArticulo_5"
        Me._lblArticulo_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_5.Size = New System.Drawing.Size(75, 13)
        Me._lblArticulo_5.TabIndex = 45
        Me._lblArticulo_5.Text = "Precio Público"
        '
        '_lblArticulo_8
        '
        Me._lblArticulo_8.AutoSize = True
        Me._lblArticulo_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_8.Location = New System.Drawing.Point(12, 124)
        Me._lblArticulo_8.Name = "_lblArticulo_8"
        Me._lblArticulo_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_8.Size = New System.Drawing.Size(78, 13)
        Me._lblArticulo_8.TabIndex = 51
        Me._lblArticulo_8.Text = "Costo Indirecto"
        '
        '_lblArticulo_7
        '
        Me._lblArticulo_7.AutoSize = True
        Me._lblArticulo_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_7.Location = New System.Drawing.Point(12, 92)
        Me._lblArticulo_7.Name = "_lblArticulo_7"
        Me._lblArticulo_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_7.Size = New System.Drawing.Size(80, 13)
        Me._lblArticulo_7.TabIndex = 49
        Me._lblArticulo_7.Text = "Costo Adicional"
        '
        '_lblArticulo_6
        '
        Me._lblArticulo_6.AutoSize = True
        Me._lblArticulo_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_6.Location = New System.Drawing.Point(12, 60)
        Me._lblArticulo_6.Name = "_lblArticulo_6"
        Me._lblArticulo_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_6.Size = New System.Drawing.Size(73, 13)
        Me._lblArticulo_6.TabIndex = 47
        Me._lblArticulo_6.Text = "Costo Factura"
        '
        '_dbcFamilia_0
        '
        Me._dbcFamilia_0.Location = New System.Drawing.Point(89, 14)
        Me._dbcFamilia_0.Name = "_dbcFamilia_0"
        Me._dbcFamilia_0.Size = New System.Drawing.Size(265, 21)
        Me._dbcFamilia_0.TabIndex = 12
        '
        '_dbcLinea_0
        '
        Me._dbcLinea_0.Location = New System.Drawing.Point(89, 46)
        Me._dbcLinea_0.Name = "_dbcLinea_0"
        Me._dbcLinea_0.Size = New System.Drawing.Size(265, 21)
        Me._dbcLinea_0.TabIndex = 14
        '
        'dbcSubLinea
        '
        Me.dbcSubLinea.Location = New System.Drawing.Point(89, 78)
        Me.dbcSubLinea.Name = "dbcSubLinea"
        Me.dbcSubLinea.Size = New System.Drawing.Size(265, 21)
        Me.dbcSubLinea.TabIndex = 16
        '
        'dbcKilates
        '
        Me.dbcKilates.Location = New System.Drawing.Point(89, 110)
        Me.dbcKilates.Name = "dbcKilates"
        Me.dbcKilates.Size = New System.Drawing.Size(134, 21)
        Me.dbcKilates.TabIndex = 18
        '
        '_dbcMaterial_0
        '
        Me._dbcMaterial_0.Location = New System.Drawing.Point(89, 140)
        Me._dbcMaterial_0.Name = "_dbcMaterial_0"
        Me._dbcMaterial_0.Size = New System.Drawing.Size(134, 21)
        Me._dbcMaterial_0.TabIndex = 20
        '
        '_lblArticulo_33
        '
        Me._lblArticulo_33.AutoSize = True
        Me._lblArticulo_33.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_33.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_33.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_33.Location = New System.Drawing.Point(2, 172)
        Me._lblArticulo_33.Name = "_lblArticulo_33"
        Me._lblArticulo_33.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_33.Size = New System.Drawing.Size(76, 13)
        Me._lblArticulo_33.TabIndex = 21
        Me._lblArticulo_33.Text = "Dato Adicional"
        '
        '_lblArticulo_9
        '
        Me._lblArticulo_9.AutoSize = True
        Me._lblArticulo_9.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_9.Location = New System.Drawing.Point(2, 144)
        Me._lblArticulo_9.Name = "_lblArticulo_9"
        Me._lblArticulo_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_9.Size = New System.Drawing.Size(83, 13)
        Me._lblArticulo_9.TabIndex = 19
        Me._lblArticulo_9.Text = "Tipo de Material"
        '
        '_lblArticulo_29
        '
        Me._lblArticulo_29.AutoSize = True
        Me._lblArticulo_29.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_29.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_29.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_29.Location = New System.Drawing.Point(0, 270)
        Me._lblArticulo_29.Name = "_lblArticulo_29"
        Me._lblArticulo_29.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_29.Size = New System.Drawing.Size(92, 13)
        Me._lblArticulo_29.TabIndex = 36
        Me._lblArticulo_29.Text = "Precio público en "
        '
        '_lblDescripcion_0
        '
        Me._lblDescripcion_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._lblDescripcion_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._lblDescripcion_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDescripcion_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
        Me._lblDescripcion_0.Location = New System.Drawing.Point(88, 230)
        Me._lblDescripcion_0.Name = "_lblDescripcion_0"
        Me._lblDescripcion_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDescripcion_0.Size = New System.Drawing.Size(537, 21)
        Me._lblDescripcion_0.TabIndex = 35
        '
        '_lblArticulo_26
        '
        Me._lblArticulo_26.AutoSize = True
        Me._lblArticulo_26.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_26.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_26.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_26.Location = New System.Drawing.Point(2, 114)
        Me._lblArticulo_26.Name = "_lblArticulo_26"
        Me._lblArticulo_26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_26.Size = New System.Drawing.Size(38, 13)
        Me._lblArticulo_26.TabIndex = 17
        Me._lblArticulo_26.Text = "Kilates"
        '
        '_lblArticulo_37
        '
        Me._lblArticulo_37.AutoSize = True
        Me._lblArticulo_37.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_37.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_37.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_37.Location = New System.Drawing.Point(313, 270)
        Me._lblArticulo_37.Name = "_lblArticulo_37"
        Me._lblArticulo_37.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_37.Size = New System.Drawing.Size(85, 13)
        Me._lblArticulo_37.TabIndex = 40
        Me._lblArticulo_37.Text = "Moneda Compra"
        '
        '_lblArticulo_4
        '
        Me._lblArticulo_4.AutoSize = True
        Me._lblArticulo_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_4.Location = New System.Drawing.Point(2, 202)
        Me._lblArticulo_4.Name = "_lblArticulo_4"
        Me._lblArticulo_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_4.Size = New System.Drawing.Size(63, 13)
        Me._lblArticulo_4.TabIndex = 33
        Me._lblArticulo_4.Text = "Descripción"
        '
        '_lblArticulo_3
        '
        Me._lblArticulo_3.AutoSize = True
        Me._lblArticulo_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_3.Location = New System.Drawing.Point(2, 82)
        Me._lblArticulo_3.Name = "_lblArticulo_3"
        Me._lblArticulo_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_3.Size = New System.Drawing.Size(54, 13)
        Me._lblArticulo_3.TabIndex = 15
        Me._lblArticulo_3.Text = "SubLínea"
        '
        '_lblArticulo_2
        '
        Me._lblArticulo_2.AutoSize = True
        Me._lblArticulo_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_2.Location = New System.Drawing.Point(2, 50)
        Me._lblArticulo_2.Name = "_lblArticulo_2"
        Me._lblArticulo_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_2.Size = New System.Drawing.Size(35, 13)
        Me._lblArticulo_2.TabIndex = 13
        Me._lblArticulo_2.Text = "Línea"
        '
        '_lblArticulo_1
        '
        Me._lblArticulo_1.AutoSize = True
        Me._lblArticulo_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_1.Location = New System.Drawing.Point(2, 18)
        Me._lblArticulo_1.Name = "_lblArticulo_1"
        Me._lblArticulo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_1.Size = New System.Drawing.Size(39, 13)
        Me._lblArticulo_1.TabIndex = 11
        Me._lblArticulo_1.Text = "Familia"
        '
        '_sstArticulo_TabPage1
        '
        Me._sstArticulo_TabPage1.Controls.Add(Me._fraContenedor_1)
        Me._sstArticulo_TabPage1.Location = New System.Drawing.Point(4, 22)
        Me._sstArticulo_TabPage1.Name = "_sstArticulo_TabPage1"
        Me._sstArticulo_TabPage1.Size = New System.Drawing.Size(700, 541)
        Me._sstArticulo_TabPage1.TabIndex = 1
        Me._sstArticulo_TabPage1.Text = "Relojería"
        '
        '_fraContenedor_1
        '
        Me._fraContenedor_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraContenedor_1.Controls.Add(Me._txtAdicional_1)
        Me._fraContenedor_1.Controls.Add(Me._fraMoneda_3)
        Me._fraContenedor_1.Controls.Add(Me._txtDescripcion_1)
        Me._fraContenedor_1.Controls.Add(Me._fraArticulo_1)
        Me._fraContenedor_1.Controls.Add(Me._fraArticulo_2)
        Me._fraContenedor_1.Controls.Add(Me._fraImagen_1)
        Me._fraContenedor_1.Controls.Add(Me._fraMoneda_1)
        Me._fraContenedor_1.Controls.Add(Me._Frame1_2)
        Me._fraContenedor_1.Controls.Add(Me._Frame2_2)
        Me._fraContenedor_1.Controls.Add(Me.chkCrono)
        Me._fraContenedor_1.Controls.Add(Me.dbcMarca)
        Me._fraContenedor_1.Controls.Add(Me.dbcModelo)
        Me._fraContenedor_1.Controls.Add(Me._dbcMaterial_1)
        Me._fraContenedor_1.Controls.Add(Me._lblArticulo_45)
        Me._fraContenedor_1.Controls.Add(Me._lblArticulo_27)
        Me._fraContenedor_1.Controls.Add(Me._lblArticulo_13)
        Me._fraContenedor_1.Controls.Add(Me._lblArticulo_12)
        Me._fraContenedor_1.Controls.Add(Me._lblArticulo_14)
        Me._fraContenedor_1.Controls.Add(Me._lblArticulo_15)
        Me._fraContenedor_1.Controls.Add(Me._lblArticulo_16)
        Me._fraContenedor_1.Controls.Add(Me._lblArticulo_17)
        Me._fraContenedor_1.Controls.Add(Me._lblArticulo_38)
        Me._fraContenedor_1.Controls.Add(Me._lblDescripcion_1)
        Me._fraContenedor_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._fraContenedor_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraContenedor_1.Location = New System.Drawing.Point(8, 24)
        Me._fraContenedor_1.Name = "_fraContenedor_1"
        Me._fraContenedor_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraContenedor_1.Size = New System.Drawing.Size(689, 481)
        Me._fraContenedor_1.TabIndex = 70
        '
        '_txtAdicional_1
        '
        Me._txtAdicional_1.AcceptsReturn = True
        Me._txtAdicional_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
        Me._txtAdicional_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtAdicional_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtAdicional_1.Location = New System.Drawing.Point(89, 168)
        Me._txtAdicional_1.MaxLength = 15
        Me._txtAdicional_1.Name = "_txtAdicional_1"
        Me._txtAdicional_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtAdicional_1.Size = New System.Drawing.Size(120, 20)
        Me._txtAdicional_1.TabIndex = 88
        '
        '_fraMoneda_3
        '
        Me._fraMoneda_3.BackColor = System.Drawing.SystemColors.Control
        Me._fraMoneda_3.Controls.Add(Me._optMoneda_7)
        Me._fraMoneda_3.Controls.Add(Me._optMoneda_6)
        Me._fraMoneda_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraMoneda_3.Location = New System.Drawing.Point(89, 256)
        Me._fraMoneda_3.Name = "_fraMoneda_3"
        Me._fraMoneda_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraMoneda_3.Size = New System.Drawing.Size(218, 33)
        Me._fraMoneda_3.TabIndex = 94
        Me._fraMoneda_3.TabStop = False
        '
        '_optMoneda_7
        '
        Me._optMoneda_7.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_7.Checked = True
        Me._optMoneda_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_7.Location = New System.Drawing.Point(36, 11)
        Me._optMoneda_7.Name = "_optMoneda_7"
        Me._optMoneda_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_7.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_7.TabIndex = 95
        Me._optMoneda_7.TabStop = True
        Me._optMoneda_7.Tag = "1"
        Me._optMoneda_7.Text = "Dólares"
        Me._optMoneda_7.UseVisualStyleBackColor = False
        '
        '_optMoneda_6
        '
        Me._optMoneda_6.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_6.Location = New System.Drawing.Point(129, 11)
        Me._optMoneda_6.Name = "_optMoneda_6"
        Me._optMoneda_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_6.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_6.TabIndex = 96
        Me._optMoneda_6.TabStop = True
        Me._optMoneda_6.Text = "Pesos"
        Me._optMoneda_6.UseVisualStyleBackColor = False
        '
        '_txtDescripcion_1
        '
        Me._txtDescripcion_1.AcceptsReturn = True
        Me._txtDescripcion_1.BackColor = System.Drawing.SystemColors.Info
        Me._txtDescripcion_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtDescripcion_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
        Me._txtDescripcion_1.Location = New System.Drawing.Point(89, 198)
        Me._txtDescripcion_1.MaxLength = 0
        Me._txtDescripcion_1.Name = "_txtDescripcion_1"
        Me._txtDescripcion_1.ReadOnly = True
        Me._txtDescripcion_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtDescripcion_1.Size = New System.Drawing.Size(537, 20)
        Me._txtDescripcion_1.TabIndex = 91
        Me.ToolTip1.SetToolTip(Me._txtDescripcion_1, "Descripción del Reloj")
        '
        '_fraArticulo_1
        '
        Me._fraArticulo_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraArticulo_1.Controls.Add(Me._optGenero_0)
        Me._fraArticulo_1.Controls.Add(Me._optGenero_1)
        Me._fraArticulo_1.Controls.Add(Me._optGenero_2)
        Me._fraArticulo_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraArticulo_1.Location = New System.Drawing.Point(89, 68)
        Me._fraArticulo_1.Name = "_fraArticulo_1"
        Me._fraArticulo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraArticulo_1.Size = New System.Drawing.Size(265, 33)
        Me._fraArticulo_1.TabIndex = 76
        Me._fraArticulo_1.TabStop = False
        '
        '_optGenero_0
        '
        Me._optGenero_0.BackColor = System.Drawing.SystemColors.Control
        Me._optGenero_0.Checked = True
        Me._optGenero_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optGenero_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optGenero_0.Location = New System.Drawing.Point(8, 8)
        Me._optGenero_0.Name = "_optGenero_0"
        Me._optGenero_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optGenero_0.Size = New System.Drawing.Size(86, 21)
        Me._optGenero_0.TabIndex = 77
        Me._optGenero_0.TabStop = True
        Me._optGenero_0.Text = "Caballero"
        Me.ToolTip1.SetToolTip(Me._optGenero_0, "Para Hombre")
        Me._optGenero_0.UseVisualStyleBackColor = False
        '
        '_optGenero_1
        '
        Me._optGenero_1.BackColor = System.Drawing.SystemColors.Control
        Me._optGenero_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optGenero_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optGenero_1.Location = New System.Drawing.Point(100, 8)
        Me._optGenero_1.Name = "_optGenero_1"
        Me._optGenero_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optGenero_1.Size = New System.Drawing.Size(57, 21)
        Me._optGenero_1.TabIndex = 78
        Me._optGenero_1.TabStop = True
        Me._optGenero_1.Text = "Dama"
        Me.ToolTip1.SetToolTip(Me._optGenero_1, "Para Mujer")
        Me._optGenero_1.UseVisualStyleBackColor = False
        '
        '_optGenero_2
        '
        Me._optGenero_2.BackColor = System.Drawing.SystemColors.Control
        Me._optGenero_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optGenero_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optGenero_2.Location = New System.Drawing.Point(190, 8)
        Me._optGenero_2.Name = "_optGenero_2"
        Me._optGenero_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optGenero_2.Size = New System.Drawing.Size(73, 21)
        Me._optGenero_2.TabIndex = 79
        Me._optGenero_2.TabStop = True
        Me._optGenero_2.Text = "Mediano"
        Me.ToolTip1.SetToolTip(Me._optGenero_2, "Para cualquier tipo de sexo")
        Me._optGenero_2.UseVisualStyleBackColor = False
        '
        '_fraArticulo_2
        '
        Me._fraArticulo_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraArticulo_2.Controls.Add(Me._optMovimiento_0)
        Me._fraArticulo_2.Controls.Add(Me._optMovimiento_1)
        Me._fraArticulo_2.Controls.Add(Me._optMovimiento_2)
        Me._fraArticulo_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraArticulo_2.Location = New System.Drawing.Point(89, 101)
        Me._fraArticulo_2.Name = "_fraArticulo_2"
        Me._fraArticulo_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraArticulo_2.Size = New System.Drawing.Size(265, 33)
        Me._fraArticulo_2.TabIndex = 81
        Me._fraArticulo_2.TabStop = False
        '
        '_optMovimiento_0
        '
        Me._optMovimiento_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMovimiento_0.Checked = True
        Me._optMovimiento_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMovimiento_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMovimiento_0.Location = New System.Drawing.Point(8, 8)
        Me._optMovimiento_0.Name = "_optMovimiento_0"
        Me._optMovimiento_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMovimiento_0.Size = New System.Drawing.Size(63, 21)
        Me._optMovimiento_0.TabIndex = 82
        Me._optMovimiento_0.TabStop = True
        Me._optMovimiento_0.Text = "Cuarzo"
        Me.ToolTip1.SetToolTip(Me._optMovimiento_0, "Movimiento por Cuarzo")
        Me._optMovimiento_0.UseVisualStyleBackColor = False
        '
        '_optMovimiento_1
        '
        Me._optMovimiento_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMovimiento_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMovimiento_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMovimiento_1.Location = New System.Drawing.Point(100, 8)
        Me._optMovimiento_1.Name = "_optMovimiento_1"
        Me._optMovimiento_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMovimiento_1.Size = New System.Drawing.Size(81, 21)
        Me._optMovimiento_1.TabIndex = 83
        Me._optMovimiento_1.TabStop = True
        Me._optMovimiento_1.Text = "Automático"
        Me.ToolTip1.SetToolTip(Me._optMovimiento_1, "Movimiento Automatizado")
        Me._optMovimiento_1.UseVisualStyleBackColor = False
        '
        '_optMovimiento_2
        '
        Me._optMovimiento_2.BackColor = System.Drawing.SystemColors.Control
        Me._optMovimiento_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMovimiento_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMovimiento_2.Location = New System.Drawing.Point(190, 8)
        Me._optMovimiento_2.Name = "_optMovimiento_2"
        Me._optMovimiento_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMovimiento_2.Size = New System.Drawing.Size(65, 21)
        Me._optMovimiento_2.TabIndex = 84
        Me._optMovimiento_2.TabStop = True
        Me._optMovimiento_2.Text = "Manual"
        Me.ToolTip1.SetToolTip(Me._optMovimiento_2, "Manual")
        Me._optMovimiento_2.UseVisualStyleBackColor = False
        '
        '_fraImagen_1
        '
        Me._fraImagen_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraImagen_1.Controls.Add(Me.Image2)
        Me._fraImagen_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraImagen_1.Location = New System.Drawing.Point(510, 8)
        Me._fraImagen_1.Name = "_fraImagen_1"
        Me._fraImagen_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraImagen_1.Size = New System.Drawing.Size(178, 186)
        Me._fraImagen_1.TabIndex = 126
        Me._fraImagen_1.TabStop = False
        Me._fraImagen_1.Text = "Imagen del Artículo"
        '
        'Image2
        '
        Me.Image2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Image2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image2.Image = Global.CorporativoV1.My.Resources.Resources.JMR3
        Me.Image2.Location = New System.Drawing.Point(7, 21)
        Me.Image2.Name = "Image2"
        Me.Image2.Size = New System.Drawing.Size(163, 157)
        Me.Image2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Image2.TabIndex = 0
        Me.Image2.TabStop = False
        '
        '_fraMoneda_1
        '
        Me._fraMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraMoneda_1.Controls.Add(Me._optMoneda_3)
        Me._fraMoneda_1.Controls.Add(Me._optMoneda_2)
        Me._fraMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraMoneda_1.Location = New System.Drawing.Point(416, 256)
        Me._fraMoneda_1.Name = "_fraMoneda_1"
        Me._fraMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraMoneda_1.Size = New System.Drawing.Size(209, 33)
        Me._fraMoneda_1.TabIndex = 98
        Me._fraMoneda_1.TabStop = False
        '
        '_optMoneda_3
        '
        Me._optMoneda_3.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_3.Location = New System.Drawing.Point(134, 11)
        Me._optMoneda_3.Name = "_optMoneda_3"
        Me._optMoneda_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_3.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_3.TabIndex = 100
        Me._optMoneda_3.TabStop = True
        Me._optMoneda_3.Text = "Pesos"
        Me._optMoneda_3.UseVisualStyleBackColor = False
        '
        '_optMoneda_2
        '
        Me._optMoneda_2.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_2.Checked = True
        Me._optMoneda_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_2.Location = New System.Drawing.Point(34, 11)
        Me._optMoneda_2.Name = "_optMoneda_2"
        Me._optMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_2.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_2.TabIndex = 99
        Me._optMoneda_2.TabStop = True
        Me._optMoneda_2.Text = "Dólares"
        Me._optMoneda_2.UseVisualStyleBackColor = False
        '
        '_Frame1_2
        '
        Me._Frame1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Frame1_2.Controls.Add(Me._txtCostoReal_1)
        Me._Frame1_2.Controls.Add(Me._txtPrecioenDolares_1)
        Me._Frame1_2.Controls.Add(Me._txtCostoIndirecto_1)
        Me._Frame1_2.Controls.Add(Me._txtCostoAdicional_1)
        Me._Frame1_2.Controls.Add(Me._txtCostoFactura_1)
        Me._Frame1_2.Controls.Add(Me._lblMargen_1)
        Me._Frame1_2.Controls.Add(Me.Label2)
        Me._Frame1_2.Controls.Add(Me._lblArticulo_40)
        Me._Frame1_2.Controls.Add(Me._lblArticulo_41)
        Me._Frame1_2.Controls.Add(Me._lblArticulo_42)
        Me._Frame1_2.Controls.Add(Me._lblArticulo_43)
        Me._Frame1_2.Controls.Add(Me._lblArticulo_44)
        Me._Frame1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame1_2.Location = New System.Drawing.Point(0, 294)
        Me._Frame1_2.Name = "_Frame1_2"
        Me._Frame1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_2.Size = New System.Drawing.Size(305, 185)
        Me._Frame1_2.TabIndex = 101
        Me._Frame1_2.TabStop = False
        '
        '_txtCostoReal_1
        '
        Me._txtCostoReal_1.AcceptsReturn = True
        Me._txtCostoReal_1.BackColor = System.Drawing.SystemColors.Info
        Me._txtCostoReal_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCostoReal_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
        Me._txtCostoReal_1.Location = New System.Drawing.Point(92, 152)
        Me._txtCostoReal_1.MaxLength = 0
        Me._txtCostoReal_1.Name = "_txtCostoReal_1"
        Me._txtCostoReal_1.ReadOnly = True
        Me._txtCostoReal_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCostoReal_1.Size = New System.Drawing.Size(113, 20)
        Me._txtCostoReal_1.TabIndex = 111
        Me._txtCostoReal_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_txtPrecioenDolares_1
        '
        Me._txtPrecioenDolares_1.AcceptsReturn = True
        Me._txtPrecioenDolares_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me._txtPrecioenDolares_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtPrecioenDolares_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._txtPrecioenDolares_1.Location = New System.Drawing.Point(92, 24)
        Me._txtPrecioenDolares_1.MaxLength = 0
        Me._txtPrecioenDolares_1.Name = "_txtPrecioenDolares_1"
        Me._txtPrecioenDolares_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtPrecioenDolares_1.Size = New System.Drawing.Size(113, 20)
        Me._txtPrecioenDolares_1.TabIndex = 103
        Me._txtPrecioenDolares_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtPrecioenDolares_1, "Precio al Público en Dólares")
        '
        '_txtCostoIndirecto_1
        '
        Me._txtCostoIndirecto_1.AcceptsReturn = True
        Me._txtCostoIndirecto_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtCostoIndirecto_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCostoIndirecto_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._txtCostoIndirecto_1.Location = New System.Drawing.Point(92, 120)
        Me._txtCostoIndirecto_1.MaxLength = 0
        Me._txtCostoIndirecto_1.Name = "_txtCostoIndirecto_1"
        Me._txtCostoIndirecto_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCostoIndirecto_1.Size = New System.Drawing.Size(113, 20)
        Me._txtCostoIndirecto_1.TabIndex = 109
        Me._txtCostoIndirecto_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtCostoIndirecto_1, "Gastos Indirectos en Dólares")
        '
        '_txtCostoAdicional_1
        '
        Me._txtCostoAdicional_1.AcceptsReturn = True
        Me._txtCostoAdicional_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtCostoAdicional_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCostoAdicional_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._txtCostoAdicional_1.Location = New System.Drawing.Point(92, 88)
        Me._txtCostoAdicional_1.MaxLength = 0
        Me._txtCostoAdicional_1.Name = "_txtCostoAdicional_1"
        Me._txtCostoAdicional_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCostoAdicional_1.Size = New System.Drawing.Size(113, 20)
        Me._txtCostoAdicional_1.TabIndex = 107
        Me._txtCostoAdicional_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtCostoAdicional_1, "Costo en Dólares")
        '
        '_txtCostoFactura_1
        '
        Me._txtCostoFactura_1.AcceptsReturn = True
        Me._txtCostoFactura_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtCostoFactura_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCostoFactura_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._txtCostoFactura_1.Location = New System.Drawing.Point(92, 56)
        Me._txtCostoFactura_1.MaxLength = 0
        Me._txtCostoFactura_1.Name = "_txtCostoFactura_1"
        Me._txtCostoFactura_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCostoFactura_1.Size = New System.Drawing.Size(113, 20)
        Me._txtCostoFactura_1.TabIndex = 105
        Me._txtCostoFactura_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtCostoFactura_1, "Costo en Pesos")
        '
        '_lblMargen_1
        '
        Me._lblMargen_1.BackColor = System.Drawing.SystemColors.Window
        Me._lblMargen_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._lblMargen_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblMargen_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblMargen_1.Location = New System.Drawing.Point(244, 152)
        Me._lblMargen_1.Name = "_lblMargen_1"
        Me._lblMargen_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblMargen_1.Size = New System.Drawing.Size(49, 21)
        Me._lblMargen_1.TabIndex = 113
        Me._lblMargen_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(232, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(61, 29)
        Me.Label2.TabIndex = 112
        Me.Label2.Text = "% Margen de Venta "
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblArticulo_40
        '
        Me._lblArticulo_40.AutoSize = True
        Me._lblArticulo_40.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_40.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_40.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_40.Location = New System.Drawing.Point(12, 156)
        Me._lblArticulo_40.Name = "_lblArticulo_40"
        Me._lblArticulo_40.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_40.Size = New System.Drawing.Size(59, 13)
        Me._lblArticulo_40.TabIndex = 110
        Me._lblArticulo_40.Text = "Costo Real"
        Me.ToolTip1.SetToolTip(Me._lblArticulo_40, "Costo Real del artículo")
        '
        '_lblArticulo_41
        '
        Me._lblArticulo_41.AutoSize = True
        Me._lblArticulo_41.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_41.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_41.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_41.Location = New System.Drawing.Point(12, 28)
        Me._lblArticulo_41.Name = "_lblArticulo_41"
        Me._lblArticulo_41.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_41.Size = New System.Drawing.Size(75, 13)
        Me._lblArticulo_41.TabIndex = 102
        Me._lblArticulo_41.Text = "Precio Público"
        '
        '_lblArticulo_42
        '
        Me._lblArticulo_42.AutoSize = True
        Me._lblArticulo_42.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_42.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_42.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_42.Location = New System.Drawing.Point(12, 124)
        Me._lblArticulo_42.Name = "_lblArticulo_42"
        Me._lblArticulo_42.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_42.Size = New System.Drawing.Size(78, 13)
        Me._lblArticulo_42.TabIndex = 108
        Me._lblArticulo_42.Text = "Costo Indirecto"
        '
        '_lblArticulo_43
        '
        Me._lblArticulo_43.AutoSize = True
        Me._lblArticulo_43.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_43.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_43.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_43.Location = New System.Drawing.Point(12, 92)
        Me._lblArticulo_43.Name = "_lblArticulo_43"
        Me._lblArticulo_43.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_43.Size = New System.Drawing.Size(80, 13)
        Me._lblArticulo_43.TabIndex = 106
        Me._lblArticulo_43.Text = "Costo Adicional"
        '
        '_lblArticulo_44
        '
        Me._lblArticulo_44.AutoSize = True
        Me._lblArticulo_44.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_44.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_44.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_44.Location = New System.Drawing.Point(12, 60)
        Me._lblArticulo_44.Name = "_lblArticulo_44"
        Me._lblArticulo_44.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_44.Size = New System.Drawing.Size(73, 13)
        Me._lblArticulo_44.TabIndex = 104
        Me._lblArticulo_44.Text = "Costo Factura"
        '
        '_Frame2_2
        '
        Me._Frame2_2.BackColor = System.Drawing.SystemColors.Control
        Me._Frame2_2.Controls.Add(Me._Frame4_1)
        Me._Frame2_2.Controls.Add(Me._txtCodigodelProveedor_1)
        Me._Frame2_2.Controls.Add(Me._dbcProveedor_1)
        Me._Frame2_2.Controls.Add(Me._cboUnidad_1)
        Me._Frame2_2.Controls.Add(Me._cboAlmacen_1)
        Me._Frame2_2.Controls.Add(Me._lblArticulo_18)
        Me._Frame2_2.Controls.Add(Me._lblArticulo_19)
        Me._Frame2_2.Controls.Add(Me._lblArticulo_20)
        Me._Frame2_2.Controls.Add(Me._lblArticulo_21)
        Me._Frame2_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame2_2.Location = New System.Drawing.Point(312, 294)
        Me._Frame2_2.Name = "_Frame2_2"
        Me._Frame2_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame2_2.Size = New System.Drawing.Size(313, 185)
        Me._Frame2_2.TabIndex = 114
        Me._Frame2_2.TabStop = False
        '
        '_Frame4_1
        '
        Me._Frame4_1.BackColor = System.Drawing.SystemColors.Control
        Me._Frame4_1.Controls.Add(Me._txtImagen_1)
        Me._Frame4_1.Controls.Add(Me._cmdBuscarImagen_1)
        Me._Frame4_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._Frame4_1.Location = New System.Drawing.Point(12, 132)
        Me._Frame4_1.Name = "_Frame4_1"
        Me._Frame4_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame4_1.Size = New System.Drawing.Size(290, 44)
        Me._Frame4_1.TabIndex = 123
        Me._Frame4_1.TabStop = False
        Me._Frame4_1.Text = "Imagen"
        '
        '_txtImagen_1
        '
        Me._txtImagen_1.AcceptsReturn = True
        Me._txtImagen_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtImagen_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtImagen_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtImagen_1.Location = New System.Drawing.Point(9, 15)
        Me._txtImagen_1.MaxLength = 0
        Me._txtImagen_1.Name = "_txtImagen_1"
        Me._txtImagen_1.ReadOnly = True
        Me._txtImagen_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtImagen_1.Size = New System.Drawing.Size(245, 20)
        Me._txtImagen_1.TabIndex = 124
        '
        '_cmdBuscarImagen_1
        '
        Me._cmdBuscarImagen_1.BackColor = System.Drawing.SystemColors.Control
        Me._cmdBuscarImagen_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdBuscarImagen_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._cmdBuscarImagen_1.Location = New System.Drawing.Point(260, 15)
        Me._cmdBuscarImagen_1.Name = "_cmdBuscarImagen_1"
        Me._cmdBuscarImagen_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdBuscarImagen_1.Size = New System.Drawing.Size(22, 21)
        Me._cmdBuscarImagen_1.TabIndex = 125
        Me._cmdBuscarImagen_1.Text = "..."
        Me._cmdBuscarImagen_1.UseVisualStyleBackColor = False
        '
        '_txtCodigodelProveedor_1
        '
        Me._txtCodigodelProveedor_1.AcceptsReturn = True
        Me._txtCodigodelProveedor_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
        Me._txtCodigodelProveedor_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCodigodelProveedor_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtCodigodelProveedor_1.Location = New System.Drawing.Point(171, 102)
        Me._txtCodigodelProveedor_1.MaxLength = 20
        Me._txtCodigodelProveedor_1.Name = "_txtCodigodelProveedor_1"
        Me._txtCodigodelProveedor_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCodigodelProveedor_1.Size = New System.Drawing.Size(129, 20)
        Me._txtCodigodelProveedor_1.TabIndex = 122
        Me.ToolTip1.SetToolTip(Me._txtCodigodelProveedor_1, "Código que usa el Proveedor para el Artículo")
        '
        '_dbcProveedor_1
        '
        Me._dbcProveedor_1.Location = New System.Drawing.Point(100, 74)
        Me._dbcProveedor_1.Name = "_dbcProveedor_1"
        Me._dbcProveedor_1.Size = New System.Drawing.Size(201, 21)
        Me._dbcProveedor_1.TabIndex = 120
        '
        '_cboUnidad_1
        '
        Me._cboUnidad_1.Location = New System.Drawing.Point(100, 17)
        Me._cboUnidad_1.Name = "_cboUnidad_1"
        Me._cboUnidad_1.Size = New System.Drawing.Size(78, 21)
        Me._cboUnidad_1.TabIndex = 116
        '
        '_cboAlmacen_1
        '
        Me._cboAlmacen_1.Location = New System.Drawing.Point(100, 46)
        Me._cboAlmacen_1.Name = "_cboAlmacen_1"
        Me._cboAlmacen_1.Size = New System.Drawing.Size(201, 21)
        Me._cboAlmacen_1.TabIndex = 118
        '
        '_lblArticulo_18
        '
        Me._lblArticulo_18.AutoSize = True
        Me._lblArticulo_18.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_18.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_18.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_18.Location = New System.Drawing.Point(12, 50)
        Me._lblArticulo_18.Name = "_lblArticulo_18"
        Me._lblArticulo_18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_18.Size = New System.Drawing.Size(84, 13)
        Me._lblArticulo_18.TabIndex = 117
        Me._lblArticulo_18.Text = "Almacén/Origen"
        '
        '_lblArticulo_19
        '
        Me._lblArticulo_19.AutoSize = True
        Me._lblArticulo_19.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_19.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_19.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_19.Location = New System.Drawing.Point(12, 21)
        Me._lblArticulo_19.Name = "_lblArticulo_19"
        Me._lblArticulo_19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_19.Size = New System.Drawing.Size(41, 13)
        Me._lblArticulo_19.TabIndex = 115
        Me._lblArticulo_19.Text = "Unidad"
        '
        '_lblArticulo_20
        '
        Me._lblArticulo_20.AutoSize = True
        Me._lblArticulo_20.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_20.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_20.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_20.Location = New System.Drawing.Point(19, 106)
        Me._lblArticulo_20.Name = "_lblArticulo_20"
        Me._lblArticulo_20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_20.Size = New System.Drawing.Size(156, 13)
        Me._lblArticulo_20.TabIndex = 121
        Me._lblArticulo_20.Text = "Código artículo del proveedor : "
        '
        '_lblArticulo_21
        '
        Me._lblArticulo_21.AutoSize = True
        Me._lblArticulo_21.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_21.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_21.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_21.Location = New System.Drawing.Point(12, 79)
        Me._lblArticulo_21.Name = "_lblArticulo_21"
        Me._lblArticulo_21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_21.Size = New System.Drawing.Size(56, 13)
        Me._lblArticulo_21.TabIndex = 119
        Me._lblArticulo_21.Text = "Proveedor"
        '
        'chkCrono
        '
        Me.chkCrono.BackColor = System.Drawing.SystemColors.Control
        Me.chkCrono.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCrono.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCrono.Location = New System.Drawing.Point(279, 174)
        Me.chkCrono.Name = "chkCrono"
        Me.chkCrono.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCrono.Size = New System.Drawing.Size(81, 17)
        Me.chkCrono.TabIndex = 89
        Me.chkCrono.Text = "Cronógrafo"
        Me.chkCrono.UseVisualStyleBackColor = False
        '
        'dbcMarca
        '
        Me.dbcMarca.Location = New System.Drawing.Point(89, 14)
        Me.dbcMarca.Name = "dbcMarca"
        Me.dbcMarca.Size = New System.Drawing.Size(265, 21)
        Me.dbcMarca.TabIndex = 72
        '
        'dbcModelo
        '
        Me.dbcModelo.Location = New System.Drawing.Point(89, 46)
        Me.dbcModelo.Name = "dbcModelo"
        Me.dbcModelo.Size = New System.Drawing.Size(265, 21)
        Me.dbcModelo.TabIndex = 74
        '
        '_dbcMaterial_1
        '
        Me._dbcMaterial_1.Location = New System.Drawing.Point(89, 140)
        Me._dbcMaterial_1.Name = "_dbcMaterial_1"
        Me._dbcMaterial_1.Size = New System.Drawing.Size(265, 21)
        Me._dbcMaterial_1.TabIndex = 86
        '
        '_lblArticulo_45
        '
        Me._lblArticulo_45.AutoSize = True
        Me._lblArticulo_45.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_45.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_45.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_45.Location = New System.Drawing.Point(2, 172)
        Me._lblArticulo_45.Name = "_lblArticulo_45"
        Me._lblArticulo_45.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_45.Size = New System.Drawing.Size(76, 13)
        Me._lblArticulo_45.TabIndex = 87
        Me._lblArticulo_45.Text = "Dato Adicional"
        '
        '_lblArticulo_27
        '
        Me._lblArticulo_27.AutoSize = True
        Me._lblArticulo_27.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_27.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_27.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_27.Location = New System.Drawing.Point(0, 270)
        Me._lblArticulo_27.Name = "_lblArticulo_27"
        Me._lblArticulo_27.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_27.Size = New System.Drawing.Size(92, 13)
        Me._lblArticulo_27.TabIndex = 93
        Me._lblArticulo_27.Text = "Precio público en "
        '
        '_lblArticulo_13
        '
        Me._lblArticulo_13.AutoSize = True
        Me._lblArticulo_13.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_13.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_13.Location = New System.Drawing.Point(2, 50)
        Me._lblArticulo_13.Name = "_lblArticulo_13"
        Me._lblArticulo_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_13.Size = New System.Drawing.Size(42, 13)
        Me._lblArticulo_13.TabIndex = 73
        Me._lblArticulo_13.Text = "Modelo"
        '
        '_lblArticulo_12
        '
        Me._lblArticulo_12.AutoSize = True
        Me._lblArticulo_12.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_12.Location = New System.Drawing.Point(2, 18)
        Me._lblArticulo_12.Name = "_lblArticulo_12"
        Me._lblArticulo_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_12.Size = New System.Drawing.Size(37, 13)
        Me._lblArticulo_12.TabIndex = 71
        Me._lblArticulo_12.Text = "Marca"
        '
        '_lblArticulo_14
        '
        Me._lblArticulo_14.AutoSize = True
        Me._lblArticulo_14.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_14.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_14.Location = New System.Drawing.Point(2, 202)
        Me._lblArticulo_14.Name = "_lblArticulo_14"
        Me._lblArticulo_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_14.Size = New System.Drawing.Size(63, 13)
        Me._lblArticulo_14.TabIndex = 90
        Me._lblArticulo_14.Text = "Descripción"
        '
        '_lblArticulo_15
        '
        Me._lblArticulo_15.AutoSize = True
        Me._lblArticulo_15.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_15.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_15.Location = New System.Drawing.Point(2, 82)
        Me._lblArticulo_15.Name = "_lblArticulo_15"
        Me._lblArticulo_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_15.Size = New System.Drawing.Size(42, 13)
        Me._lblArticulo_15.TabIndex = 75
        Me._lblArticulo_15.Text = "Género"
        '
        '_lblArticulo_16
        '
        Me._lblArticulo_16.AutoSize = True
        Me._lblArticulo_16.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_16.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_16.Location = New System.Drawing.Point(2, 114)
        Me._lblArticulo_16.Name = "_lblArticulo_16"
        Me._lblArticulo_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_16.Size = New System.Drawing.Size(82, 13)
        Me._lblArticulo_16.TabIndex = 80
        Me._lblArticulo_16.Text = "Funcionamiento"
        '
        '_lblArticulo_17
        '
        Me._lblArticulo_17.AutoSize = True
        Me._lblArticulo_17.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_17.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_17.Location = New System.Drawing.Point(2, 144)
        Me._lblArticulo_17.Name = "_lblArticulo_17"
        Me._lblArticulo_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_17.Size = New System.Drawing.Size(83, 13)
        Me._lblArticulo_17.TabIndex = 85
        Me._lblArticulo_17.Text = "Tipo de Material"
        '
        '_lblArticulo_38
        '
        Me._lblArticulo_38.AutoSize = True
        Me._lblArticulo_38.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_38.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_38.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_38.Location = New System.Drawing.Point(313, 270)
        Me._lblArticulo_38.Name = "_lblArticulo_38"
        Me._lblArticulo_38.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_38.Size = New System.Drawing.Size(85, 13)
        Me._lblArticulo_38.TabIndex = 97
        Me._lblArticulo_38.Text = "Moneda Compra"
        '
        '_lblDescripcion_1
        '
        Me._lblDescripcion_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._lblDescripcion_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._lblDescripcion_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDescripcion_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
        Me._lblDescripcion_1.Location = New System.Drawing.Point(89, 230)
        Me._lblDescripcion_1.Name = "_lblDescripcion_1"
        Me._lblDescripcion_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDescripcion_1.Size = New System.Drawing.Size(537, 21)
        Me._lblDescripcion_1.TabIndex = 92
        '
        '_sstArticulo_TabPage2
        '
        Me._sstArticulo_TabPage2.Controls.Add(Me._fraContenedor_2)
        Me._sstArticulo_TabPage2.Location = New System.Drawing.Point(4, 22)
        Me._sstArticulo_TabPage2.Name = "_sstArticulo_TabPage2"
        Me._sstArticulo_TabPage2.Size = New System.Drawing.Size(700, 541)
        Me._sstArticulo_TabPage2.TabIndex = 2
        Me._sstArticulo_TabPage2.Text = "Varios"
        '
        '_fraContenedor_2
        '
        Me._fraContenedor_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraContenedor_2.Controls.Add(Me._txtAdicional_2)
        Me._fraContenedor_2.Controls.Add(Me._fraMoneda_4)
        Me._fraContenedor_2.Controls.Add(Me._fraMoneda_2)
        Me._fraContenedor_2.Controls.Add(Me._Frame2_3)
        Me._fraContenedor_2.Controls.Add(Me._Frame1_3)
        Me._fraContenedor_2.Controls.Add(Me._txtDescripcion_2)
        Me._fraContenedor_2.Controls.Add(Me._fraImagen_2)
        Me._fraContenedor_2.Controls.Add(Me._dbcFamilia_1)
        Me._fraContenedor_2.Controls.Add(Me._dbcLinea_1)
        Me._fraContenedor_2.Controls.Add(Me._dbcMaterial_2)
        Me._fraContenedor_2.Controls.Add(Me._lblArticulo_54)
        Me._fraContenedor_2.Controls.Add(Me._lblArticulo_49)
        Me._fraContenedor_2.Controls.Add(Me._lblArticulo_28)
        Me._fraContenedor_2.Controls.Add(Me._lblArticulo_24)
        Me._fraContenedor_2.Controls.Add(Me._lblArticulo_25)
        Me._fraContenedor_2.Controls.Add(Me._lblArticulo_30)
        Me._fraContenedor_2.Controls.Add(Me._lblArticulo_39)
        Me._fraContenedor_2.Controls.Add(Me._lblDescripcion_2)
        Me._fraContenedor_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._fraContenedor_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraContenedor_2.Location = New System.Drawing.Point(8, 24)
        Me._fraContenedor_2.Name = "_fraContenedor_2"
        Me._fraContenedor_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraContenedor_2.Size = New System.Drawing.Size(689, 481)
        Me._fraContenedor_2.TabIndex = 127
        '
        '_txtAdicional_2
        '
        Me._txtAdicional_2.AcceptsReturn = True
        Me._txtAdicional_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
        Me._txtAdicional_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtAdicional_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtAdicional_2.Location = New System.Drawing.Point(89, 110)
        Me._txtAdicional_2.MaxLength = 15
        Me._txtAdicional_2.Name = "_txtAdicional_2"
        Me._txtAdicional_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtAdicional_2.Size = New System.Drawing.Size(120, 20)
        Me._txtAdicional_2.TabIndex = 135
        '
        '_fraMoneda_4
        '
        Me._fraMoneda_4.BackColor = System.Drawing.SystemColors.Control
        Me._fraMoneda_4.Controls.Add(Me._optMoneda_9)
        Me._fraMoneda_4.Controls.Add(Me._optMoneda_8)
        Me._fraMoneda_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraMoneda_4.Location = New System.Drawing.Point(89, 256)
        Me._fraMoneda_4.Name = "_fraMoneda_4"
        Me._fraMoneda_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraMoneda_4.Size = New System.Drawing.Size(218, 33)
        Me._fraMoneda_4.TabIndex = 140
        Me._fraMoneda_4.TabStop = False
        '
        '_optMoneda_9
        '
        Me._optMoneda_9.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_9.Location = New System.Drawing.Point(129, 11)
        Me._optMoneda_9.Name = "_optMoneda_9"
        Me._optMoneda_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_9.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_9.TabIndex = 142
        Me._optMoneda_9.TabStop = True
        Me._optMoneda_9.Text = "Pesos"
        Me._optMoneda_9.UseVisualStyleBackColor = False
        '
        '_optMoneda_8
        '
        Me._optMoneda_8.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_8.Checked = True
        Me._optMoneda_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_8.Location = New System.Drawing.Point(36, 11)
        Me._optMoneda_8.Name = "_optMoneda_8"
        Me._optMoneda_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_8.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_8.TabIndex = 141
        Me._optMoneda_8.TabStop = True
        Me._optMoneda_8.Tag = "1"
        Me._optMoneda_8.Text = "Dólares"
        Me._optMoneda_8.UseVisualStyleBackColor = False
        '
        '_fraMoneda_2
        '
        Me._fraMoneda_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraMoneda_2.Controls.Add(Me._optMoneda_4)
        Me._fraMoneda_2.Controls.Add(Me._optMoneda_5)
        Me._fraMoneda_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraMoneda_2.Location = New System.Drawing.Point(416, 256)
        Me._fraMoneda_2.Name = "_fraMoneda_2"
        Me._fraMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraMoneda_2.Size = New System.Drawing.Size(209, 33)
        Me._fraMoneda_2.TabIndex = 144
        Me._fraMoneda_2.TabStop = False
        '
        '_optMoneda_4
        '
        Me._optMoneda_4.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_4.Checked = True
        Me._optMoneda_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_4.Location = New System.Drawing.Point(34, 11)
        Me._optMoneda_4.Name = "_optMoneda_4"
        Me._optMoneda_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_4.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_4.TabIndex = 145
        Me._optMoneda_4.TabStop = True
        Me._optMoneda_4.Text = "Dólares"
        Me._optMoneda_4.UseVisualStyleBackColor = False
        '
        '_optMoneda_5
        '
        Me._optMoneda_5.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optMoneda_5.Location = New System.Drawing.Point(134, 11)
        Me._optMoneda_5.Name = "_optMoneda_5"
        Me._optMoneda_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_5.Size = New System.Drawing.Size(65, 17)
        Me._optMoneda_5.TabIndex = 146
        Me._optMoneda_5.TabStop = True
        Me._optMoneda_5.Text = "Pesos"
        Me._optMoneda_5.UseVisualStyleBackColor = False
        '
        '_Frame2_3
        '
        Me._Frame2_3.BackColor = System.Drawing.SystemColors.Control
        Me._Frame2_3.Controls.Add(Me._Frame4_2)
        Me._Frame2_3.Controls.Add(Me._txtCodigodelProveedor_2)
        Me._Frame2_3.Controls.Add(Me._dbcProveedor_2)
        Me._Frame2_3.Controls.Add(Me._cboUnidad_2)
        Me._Frame2_3.Controls.Add(Me._cboAlmacen_2)
        Me._Frame2_3.Controls.Add(Me._lblArticulo_50)
        Me._Frame2_3.Controls.Add(Me._lblArticulo_51)
        Me._Frame2_3.Controls.Add(Me._lblArticulo_52)
        Me._Frame2_3.Controls.Add(Me._lblArticulo_53)
        Me._Frame2_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame2_3.Location = New System.Drawing.Point(312, 294)
        Me._Frame2_3.Name = "_Frame2_3"
        Me._Frame2_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame2_3.Size = New System.Drawing.Size(313, 185)
        Me._Frame2_3.TabIndex = 160
        Me._Frame2_3.TabStop = False
        '
        '_Frame4_2
        '
        Me._Frame4_2.BackColor = System.Drawing.SystemColors.Control
        Me._Frame4_2.Controls.Add(Me._txtImagen_2)
        Me._Frame4_2.Controls.Add(Me._cmdBuscarImagen_2)
        Me._Frame4_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._Frame4_2.Location = New System.Drawing.Point(12, 132)
        Me._Frame4_2.Name = "_Frame4_2"
        Me._Frame4_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame4_2.Size = New System.Drawing.Size(290, 44)
        Me._Frame4_2.TabIndex = 169
        Me._Frame4_2.TabStop = False
        Me._Frame4_2.Text = "Imagen"
        '
        '_txtImagen_2
        '
        Me._txtImagen_2.AcceptsReturn = True
        Me._txtImagen_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtImagen_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtImagen_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtImagen_2.Location = New System.Drawing.Point(9, 15)
        Me._txtImagen_2.MaxLength = 0
        Me._txtImagen_2.Name = "_txtImagen_2"
        Me._txtImagen_2.ReadOnly = True
        Me._txtImagen_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtImagen_2.Size = New System.Drawing.Size(245, 20)
        Me._txtImagen_2.TabIndex = 170
        '
        '_cmdBuscarImagen_2
        '
        Me._cmdBuscarImagen_2.BackColor = System.Drawing.SystemColors.Control
        Me._cmdBuscarImagen_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdBuscarImagen_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._cmdBuscarImagen_2.Location = New System.Drawing.Point(260, 15)
        Me._cmdBuscarImagen_2.Name = "_cmdBuscarImagen_2"
        Me._cmdBuscarImagen_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdBuscarImagen_2.Size = New System.Drawing.Size(22, 21)
        Me._cmdBuscarImagen_2.TabIndex = 171
        Me._cmdBuscarImagen_2.Text = "..."
        Me._cmdBuscarImagen_2.UseVisualStyleBackColor = False
        '
        '_txtCodigodelProveedor_2
        '
        Me._txtCodigodelProveedor_2.AcceptsReturn = True
        Me._txtCodigodelProveedor_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(230, Byte), Integer), CType(CType(244, Byte), Integer))
        Me._txtCodigodelProveedor_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCodigodelProveedor_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtCodigodelProveedor_2.Location = New System.Drawing.Point(171, 102)
        Me._txtCodigodelProveedor_2.MaxLength = 20
        Me._txtCodigodelProveedor_2.Name = "_txtCodigodelProveedor_2"
        Me._txtCodigodelProveedor_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCodigodelProveedor_2.Size = New System.Drawing.Size(129, 20)
        Me._txtCodigodelProveedor_2.TabIndex = 168
        Me.ToolTip1.SetToolTip(Me._txtCodigodelProveedor_2, "Código que usa el Proveedor para el Artículo")
        '
        '_dbcProveedor_2
        '
        Me._dbcProveedor_2.Location = New System.Drawing.Point(100, 74)
        Me._dbcProveedor_2.Name = "_dbcProveedor_2"
        Me._dbcProveedor_2.Size = New System.Drawing.Size(201, 21)
        Me._dbcProveedor_2.TabIndex = 166
        '
        '_cboUnidad_2
        '
        Me._cboUnidad_2.Location = New System.Drawing.Point(100, 17)
        Me._cboUnidad_2.Name = "_cboUnidad_2"
        Me._cboUnidad_2.Size = New System.Drawing.Size(78, 21)
        Me._cboUnidad_2.TabIndex = 162
        '
        '_cboAlmacen_2
        '
        Me._cboAlmacen_2.Location = New System.Drawing.Point(100, 46)
        Me._cboAlmacen_2.Name = "_cboAlmacen_2"
        Me._cboAlmacen_2.Size = New System.Drawing.Size(201, 21)
        Me._cboAlmacen_2.TabIndex = 164
        '
        '_lblArticulo_50
        '
        Me._lblArticulo_50.AutoSize = True
        Me._lblArticulo_50.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_50.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_50.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_50.Location = New System.Drawing.Point(12, 79)
        Me._lblArticulo_50.Name = "_lblArticulo_50"
        Me._lblArticulo_50.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_50.Size = New System.Drawing.Size(56, 13)
        Me._lblArticulo_50.TabIndex = 165
        Me._lblArticulo_50.Text = "Proveedor"
        '
        '_lblArticulo_51
        '
        Me._lblArticulo_51.AutoSize = True
        Me._lblArticulo_51.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_51.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_51.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_51.Location = New System.Drawing.Point(19, 106)
        Me._lblArticulo_51.Name = "_lblArticulo_51"
        Me._lblArticulo_51.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_51.Size = New System.Drawing.Size(156, 13)
        Me._lblArticulo_51.TabIndex = 167
        Me._lblArticulo_51.Text = "Código artículo del proveedor : "
        '
        '_lblArticulo_52
        '
        Me._lblArticulo_52.AutoSize = True
        Me._lblArticulo_52.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_52.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_52.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_52.Location = New System.Drawing.Point(12, 21)
        Me._lblArticulo_52.Name = "_lblArticulo_52"
        Me._lblArticulo_52.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_52.Size = New System.Drawing.Size(41, 13)
        Me._lblArticulo_52.TabIndex = 161
        Me._lblArticulo_52.Text = "Unidad"
        '
        '_lblArticulo_53
        '
        Me._lblArticulo_53.AutoSize = True
        Me._lblArticulo_53.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_53.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_53.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_53.Location = New System.Drawing.Point(12, 50)
        Me._lblArticulo_53.Name = "_lblArticulo_53"
        Me._lblArticulo_53.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_53.Size = New System.Drawing.Size(84, 13)
        Me._lblArticulo_53.TabIndex = 163
        Me._lblArticulo_53.Text = "Almacén/Origen"
        '
        '_Frame1_3
        '
        Me._Frame1_3.BackColor = System.Drawing.SystemColors.Control
        Me._Frame1_3.Controls.Add(Me._txtCostoFactura_2)
        Me._Frame1_3.Controls.Add(Me._txtCostoAdicional_2)
        Me._Frame1_3.Controls.Add(Me._txtCostoIndirecto_2)
        Me._Frame1_3.Controls.Add(Me._txtPrecioenDolares_2)
        Me._Frame1_3.Controls.Add(Me._txtCostoReal_2)
        Me._Frame1_3.Controls.Add(Me._lblMargen_2)
        Me._Frame1_3.Controls.Add(Me.Label3)
        Me._Frame1_3.Controls.Add(Me._lblArticulo_22)
        Me._Frame1_3.Controls.Add(Me._lblArticulo_23)
        Me._Frame1_3.Controls.Add(Me._lblArticulo_46)
        Me._Frame1_3.Controls.Add(Me._lblArticulo_47)
        Me._Frame1_3.Controls.Add(Me._lblArticulo_48)
        Me._Frame1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame1_3.Location = New System.Drawing.Point(0, 294)
        Me._Frame1_3.Name = "_Frame1_3"
        Me._Frame1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_3.Size = New System.Drawing.Size(305, 185)
        Me._Frame1_3.TabIndex = 147
        Me._Frame1_3.TabStop = False
        '
        '_txtCostoFactura_2
        '
        Me._txtCostoFactura_2.AcceptsReturn = True
        Me._txtCostoFactura_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtCostoFactura_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCostoFactura_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._txtCostoFactura_2.Location = New System.Drawing.Point(92, 56)
        Me._txtCostoFactura_2.MaxLength = 0
        Me._txtCostoFactura_2.Name = "_txtCostoFactura_2"
        Me._txtCostoFactura_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCostoFactura_2.Size = New System.Drawing.Size(113, 20)
        Me._txtCostoFactura_2.TabIndex = 151
        Me._txtCostoFactura_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtCostoFactura_2, "Costo en Pesos")
        '
        '_txtCostoAdicional_2
        '
        Me._txtCostoAdicional_2.AcceptsReturn = True
        Me._txtCostoAdicional_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtCostoAdicional_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCostoAdicional_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._txtCostoAdicional_2.Location = New System.Drawing.Point(92, 88)
        Me._txtCostoAdicional_2.MaxLength = 0
        Me._txtCostoAdicional_2.Name = "_txtCostoAdicional_2"
        Me._txtCostoAdicional_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCostoAdicional_2.Size = New System.Drawing.Size(113, 20)
        Me._txtCostoAdicional_2.TabIndex = 153
        Me._txtCostoAdicional_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtCostoAdicional_2, "Costo en Dólares")
        '
        '_txtCostoIndirecto_2
        '
        Me._txtCostoIndirecto_2.AcceptsReturn = True
        Me._txtCostoIndirecto_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtCostoIndirecto_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCostoIndirecto_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._txtCostoIndirecto_2.Location = New System.Drawing.Point(92, 120)
        Me._txtCostoIndirecto_2.MaxLength = 0
        Me._txtCostoIndirecto_2.Name = "_txtCostoIndirecto_2"
        Me._txtCostoIndirecto_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCostoIndirecto_2.Size = New System.Drawing.Size(113, 20)
        Me._txtCostoIndirecto_2.TabIndex = 155
        Me._txtCostoIndirecto_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtCostoIndirecto_2, "Gastos Indirectos en Dólares")
        '
        '_txtPrecioenDolares_2
        '
        Me._txtPrecioenDolares_2.AcceptsReturn = True
        Me._txtPrecioenDolares_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(213, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(213, Byte), Integer))
        Me._txtPrecioenDolares_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtPrecioenDolares_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtPrecioenDolares_2.Location = New System.Drawing.Point(92, 24)
        Me._txtPrecioenDolares_2.MaxLength = 0
        Me._txtPrecioenDolares_2.Name = "_txtPrecioenDolares_2"
        Me._txtPrecioenDolares_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtPrecioenDolares_2.Size = New System.Drawing.Size(113, 20)
        Me._txtPrecioenDolares_2.TabIndex = 149
        Me._txtPrecioenDolares_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtPrecioenDolares_2, "Precio al Público en Dólares")
        '
        '_txtCostoReal_2
        '
        Me._txtCostoReal_2.AcceptsReturn = True
        Me._txtCostoReal_2.BackColor = System.Drawing.SystemColors.Info
        Me._txtCostoReal_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtCostoReal_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
        Me._txtCostoReal_2.Location = New System.Drawing.Point(92, 152)
        Me._txtCostoReal_2.MaxLength = 0
        Me._txtCostoReal_2.Name = "_txtCostoReal_2"
        Me._txtCostoReal_2.ReadOnly = True
        Me._txtCostoReal_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtCostoReal_2.Size = New System.Drawing.Size(113, 20)
        Me._txtCostoReal_2.TabIndex = 157
        Me._txtCostoReal_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me._txtCostoReal_2, "Costo Real del artículo")
        '
        '_lblMargen_2
        '
        Me._lblMargen_2.BackColor = System.Drawing.SystemColors.Window
        Me._lblMargen_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._lblMargen_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblMargen_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblMargen_2.Location = New System.Drawing.Point(244, 152)
        Me._lblMargen_2.Name = "_lblMargen_2"
        Me._lblMargen_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblMargen_2.Size = New System.Drawing.Size(49, 21)
        Me._lblMargen_2.TabIndex = 159
        Me._lblMargen_2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(232, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(61, 29)
        Me.Label3.TabIndex = 158
        Me.Label3.Text = "% Margen de Venta "
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_lblArticulo_22
        '
        Me._lblArticulo_22.AutoSize = True
        Me._lblArticulo_22.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_22.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_22.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_22.Location = New System.Drawing.Point(12, 60)
        Me._lblArticulo_22.Name = "_lblArticulo_22"
        Me._lblArticulo_22.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_22.Size = New System.Drawing.Size(73, 13)
        Me._lblArticulo_22.TabIndex = 150
        Me._lblArticulo_22.Text = "Costo Factura"
        '
        '_lblArticulo_23
        '
        Me._lblArticulo_23.AutoSize = True
        Me._lblArticulo_23.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_23.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_23.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_23.Location = New System.Drawing.Point(12, 92)
        Me._lblArticulo_23.Name = "_lblArticulo_23"
        Me._lblArticulo_23.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_23.Size = New System.Drawing.Size(80, 13)
        Me._lblArticulo_23.TabIndex = 152
        Me._lblArticulo_23.Text = "Costo Adicional"
        '
        '_lblArticulo_46
        '
        Me._lblArticulo_46.AutoSize = True
        Me._lblArticulo_46.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_46.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_46.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_46.Location = New System.Drawing.Point(12, 124)
        Me._lblArticulo_46.Name = "_lblArticulo_46"
        Me._lblArticulo_46.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_46.Size = New System.Drawing.Size(78, 13)
        Me._lblArticulo_46.TabIndex = 154
        Me._lblArticulo_46.Text = "Costo Indirecto"
        '
        '_lblArticulo_47
        '
        Me._lblArticulo_47.AutoSize = True
        Me._lblArticulo_47.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_47.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_47.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_47.Location = New System.Drawing.Point(12, 28)
        Me._lblArticulo_47.Name = "_lblArticulo_47"
        Me._lblArticulo_47.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_47.Size = New System.Drawing.Size(78, 13)
        Me._lblArticulo_47.TabIndex = 148
        Me._lblArticulo_47.Text = "Precio Público "
        '
        '_lblArticulo_48
        '
        Me._lblArticulo_48.AutoSize = True
        Me._lblArticulo_48.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_48.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_48.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_48.Location = New System.Drawing.Point(12, 156)
        Me._lblArticulo_48.Name = "_lblArticulo_48"
        Me._lblArticulo_48.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_48.Size = New System.Drawing.Size(59, 13)
        Me._lblArticulo_48.TabIndex = 156
        Me._lblArticulo_48.Text = "Costo Real"
        '
        '_txtDescripcion_2
        '
        Me._txtDescripcion_2.AcceptsReturn = True
        Me._txtDescripcion_2.BackColor = System.Drawing.SystemColors.Info
        Me._txtDescripcion_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtDescripcion_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
        Me._txtDescripcion_2.Location = New System.Drawing.Point(89, 198)
        Me._txtDescripcion_2.MaxLength = 0
        Me._txtDescripcion_2.Name = "_txtDescripcion_2"
        Me._txtDescripcion_2.ReadOnly = True
        Me._txtDescripcion_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtDescripcion_2.Size = New System.Drawing.Size(537, 20)
        Me._txtDescripcion_2.TabIndex = 137
        Me.ToolTip1.SetToolTip(Me._txtDescripcion_2, "Descripción")
        '
        '_fraImagen_2
        '
        Me._fraImagen_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraImagen_2.Controls.Add(Me.Image3)
        Me._fraImagen_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraImagen_2.Location = New System.Drawing.Point(510, 8)
        Me._fraImagen_2.Name = "_fraImagen_2"
        Me._fraImagen_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraImagen_2.Size = New System.Drawing.Size(178, 186)
        Me._fraImagen_2.TabIndex = 172
        Me._fraImagen_2.TabStop = False
        Me._fraImagen_2.Text = "Imagen del Artículo"
        '
        'Image3
        '
        Me.Image3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Image3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image3.Image = Global.CorporativoV1.My.Resources.Resources.JMR4
        Me.Image3.Location = New System.Drawing.Point(7, 21)
        Me.Image3.Name = "Image3"
        Me.Image3.Size = New System.Drawing.Size(163, 157)
        Me.Image3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Image3.TabIndex = 0
        Me.Image3.TabStop = False
        '
        '_dbcFamilia_1
        '
        Me._dbcFamilia_1.Location = New System.Drawing.Point(89, 14)
        Me._dbcFamilia_1.Name = "_dbcFamilia_1"
        Me._dbcFamilia_1.Size = New System.Drawing.Size(265, 21)
        Me._dbcFamilia_1.TabIndex = 129
        '
        '_dbcLinea_1
        '
        Me._dbcLinea_1.Location = New System.Drawing.Point(89, 46)
        Me._dbcLinea_1.Name = "_dbcLinea_1"
        Me._dbcLinea_1.Size = New System.Drawing.Size(265, 21)
        Me._dbcLinea_1.TabIndex = 131
        '
        '_dbcMaterial_2
        '
        Me._dbcMaterial_2.Location = New System.Drawing.Point(89, 78)
        Me._dbcMaterial_2.Name = "_dbcMaterial_2"
        Me._dbcMaterial_2.Size = New System.Drawing.Size(265, 21)
        Me._dbcMaterial_2.TabIndex = 133
        '
        '_lblArticulo_54
        '
        Me._lblArticulo_54.AutoSize = True
        Me._lblArticulo_54.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_54.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_54.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_54.Location = New System.Drawing.Point(2, 114)
        Me._lblArticulo_54.Name = "_lblArticulo_54"
        Me._lblArticulo_54.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_54.Size = New System.Drawing.Size(76, 13)
        Me._lblArticulo_54.TabIndex = 134
        Me._lblArticulo_54.Text = "Dato Adicional"
        '
        '_lblArticulo_49
        '
        Me._lblArticulo_49.AutoSize = True
        Me._lblArticulo_49.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_49.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_49.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_49.Location = New System.Drawing.Point(2, 82)
        Me._lblArticulo_49.Name = "_lblArticulo_49"
        Me._lblArticulo_49.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_49.Size = New System.Drawing.Size(83, 13)
        Me._lblArticulo_49.TabIndex = 132
        Me._lblArticulo_49.Text = "Tipo de Material"
        '
        '_lblArticulo_28
        '
        Me._lblArticulo_28.AutoSize = True
        Me._lblArticulo_28.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_28.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_28.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_28.Location = New System.Drawing.Point(0, 270)
        Me._lblArticulo_28.Name = "_lblArticulo_28"
        Me._lblArticulo_28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_28.Size = New System.Drawing.Size(92, 13)
        Me._lblArticulo_28.TabIndex = 139
        Me._lblArticulo_28.Text = "Precio público en "
        '
        '_lblArticulo_24
        '
        Me._lblArticulo_24.AutoSize = True
        Me._lblArticulo_24.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_24.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_24.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_24.Location = New System.Drawing.Point(2, 18)
        Me._lblArticulo_24.Name = "_lblArticulo_24"
        Me._lblArticulo_24.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_24.Size = New System.Drawing.Size(39, 13)
        Me._lblArticulo_24.TabIndex = 128
        Me._lblArticulo_24.Text = "Familia"
        '
        '_lblArticulo_25
        '
        Me._lblArticulo_25.AutoSize = True
        Me._lblArticulo_25.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_25.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_25.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_25.Location = New System.Drawing.Point(2, 50)
        Me._lblArticulo_25.Name = "_lblArticulo_25"
        Me._lblArticulo_25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_25.Size = New System.Drawing.Size(35, 13)
        Me._lblArticulo_25.TabIndex = 130
        Me._lblArticulo_25.Text = "Línea"
        '
        '_lblArticulo_30
        '
        Me._lblArticulo_30.AutoSize = True
        Me._lblArticulo_30.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_30.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_30.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_30.Location = New System.Drawing.Point(2, 202)
        Me._lblArticulo_30.Name = "_lblArticulo_30"
        Me._lblArticulo_30.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_30.Size = New System.Drawing.Size(63, 13)
        Me._lblArticulo_30.TabIndex = 136
        Me._lblArticulo_30.Text = "Descripción"
        '
        '_lblArticulo_39
        '
        Me._lblArticulo_39.AutoSize = True
        Me._lblArticulo_39.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_39.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_39.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_39.Location = New System.Drawing.Point(313, 270)
        Me._lblArticulo_39.Name = "_lblArticulo_39"
        Me._lblArticulo_39.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_39.Size = New System.Drawing.Size(85, 13)
        Me._lblArticulo_39.TabIndex = 143
        Me._lblArticulo_39.Text = "Moneda Compra"
        '
        '_lblDescripcion_2
        '
        Me._lblDescripcion_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._lblDescripcion_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me._lblDescripcion_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblDescripcion_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(102, Byte), Integer), CType(CType(153, Byte), Integer))
        Me._lblDescripcion_2.Location = New System.Drawing.Point(89, 230)
        Me._lblDescripcion_2.Name = "_lblDescripcion_2"
        Me._lblDescripcion_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblDescripcion_2.Size = New System.Drawing.Size(537, 21)
        Me._lblDescripcion_2.TabIndex = 138
        '
        'chkCodigoAnterior
        '
        Me.chkCodigoAnterior.BackColor = System.Drawing.SystemColors.Control
        Me.chkCodigoAnterior.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCodigoAnterior.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCodigoAnterior.Location = New System.Drawing.Point(6, 0)
        Me.chkCodigoAnterior.Name = "chkCodigoAnterior"
        Me.chkCodigoAnterior.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCodigoAnterior.Size = New System.Drawing.Size(17, 17)
        Me.chkCodigoAnterior.TabIndex = 4
        Me.chkCodigoAnterior.Text = "chkCodAnterior"
        Me.chkCodigoAnterior.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.chkCodigoAnterior)
        Me.Frame3.Controls.Add(Me.txtCodArtAnterior)
        Me.Frame3.Controls.Add(Me.dbcOrigen)
        Me.Frame3.Controls.Add(Me._lblArticulo_32)
        Me.Frame3.Controls.Add(Me._lblArticulo_31)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(538, 3)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(182, 69)
        Me.Frame3.TabIndex = 3
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "      Código anterior "
        '
        'txtCodArtAnterior
        '
        Me.txtCodArtAnterior.AcceptsReturn = True
        Me.txtCodArtAnterior.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodArtAnterior.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodArtAnterior.Enabled = False
        Me.txtCodArtAnterior.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodArtAnterior.Location = New System.Drawing.Point(76, 41)
        Me.txtCodArtAnterior.MaxLength = 5
        Me.txtCodArtAnterior.Name = "txtCodArtAnterior"
        Me.txtCodArtAnterior.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodArtAnterior.Size = New System.Drawing.Size(88, 20)
        Me.txtCodArtAnterior.TabIndex = 8
        Me.txtCodArtAnterior.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'dbcOrigen
        '
        Me.dbcOrigen.Location = New System.Drawing.Point(76, 17)
        Me.dbcOrigen.Name = "dbcOrigen"
        Me.dbcOrigen.Size = New System.Drawing.Size(88, 21)
        Me.dbcOrigen.TabIndex = 6
        '
        '_lblArticulo_32
        '
        Me._lblArticulo_32.AutoSize = True
        Me._lblArticulo_32.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_32.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_32.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_32.Location = New System.Drawing.Point(31, 44)
        Me._lblArticulo_32.Name = "_lblArticulo_32"
        Me._lblArticulo_32.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_32.Size = New System.Drawing.Size(46, 13)
        Me._lblArticulo_32.TabIndex = 7
        Me._lblArticulo_32.Text = "Código: "
        '
        '_lblArticulo_31
        '
        Me._lblArticulo_31.AutoSize = True
        Me._lblArticulo_31.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_31.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_31.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblArticulo_31.Location = New System.Drawing.Point(30, 21)
        Me._lblArticulo_31.Name = "_lblArticulo_31"
        Me._lblArticulo_31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_31.Size = New System.Drawing.Size(47, 13)
        Me._lblArticulo_31.TabIndex = 5
        Me._lblArticulo_31.Text = "Origen : "
        '
        'txtCodArticulo
        '
        Me.txtCodArticulo.AcceptsReturn = True
        Me.txtCodArticulo.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodArticulo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodArticulo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodArticulo.Location = New System.Drawing.Point(72, 8)
        Me.txtCodArticulo.MaxLength = 8
        Me.txtCodArticulo.Name = "txtCodArticulo"
        Me.txtCodArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodArticulo.Size = New System.Drawing.Size(89, 20)
        Me.txtCodArticulo.TabIndex = 1
        Me.txtCodArticulo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_lblArticulo_0
        '
        Me._lblArticulo_0.AutoSize = True
        Me._lblArticulo_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblArticulo_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblArticulo_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblArticulo_0.Location = New System.Drawing.Point(16, 12)
        Me._lblArticulo_0.Name = "_lblArticulo_0"
        Me._lblArticulo_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblArticulo_0.Size = New System.Drawing.Size(44, 13)
        Me._lblArticulo_0.TabIndex = 0
        Me._lblArticulo_0.Text = "Artículo"
        '
        'cboAlmacen
        '
        Me.cboAlmacen.Location = New System.Drawing.Point(0, 0)
        Me.cboAlmacen.Name = "cboAlmacen"
        Me.cboAlmacen.Size = New System.Drawing.Size(121, 21)
        Me.cboAlmacen.TabIndex = 0
        '
        'cboUnidad
        '
        Me.cboUnidad.Location = New System.Drawing.Point(0, 0)
        Me.cboUnidad.Name = "cboUnidad"
        Me.cboUnidad.Size = New System.Drawing.Size(121, 21)
        Me.cboUnidad.TabIndex = 0
        '
        'cmdBuscarImagen
        '
        '
        'dbcFamilia
        '
        Me.dbcFamilia.Location = New System.Drawing.Point(0, 0)
        Me.dbcFamilia.Name = "dbcFamilia"
        Me.dbcFamilia.Size = New System.Drawing.Size(121, 21)
        Me.dbcFamilia.TabIndex = 0
        '
        'dbcLinea
        '
        Me.dbcLinea.Location = New System.Drawing.Point(0, 0)
        Me.dbcLinea.Name = "dbcLinea"
        Me.dbcLinea.Size = New System.Drawing.Size(121, 21)
        Me.dbcLinea.TabIndex = 0
        '
        'dbcMaterial
        '
        Me.dbcMaterial.Location = New System.Drawing.Point(0, 0)
        Me.dbcMaterial.Name = "dbcMaterial"
        Me.dbcMaterial.Size = New System.Drawing.Size(121, 21)
        Me.dbcMaterial.TabIndex = 0
        '
        'dbcProveedor
        '
        Me.dbcProveedor.Location = New System.Drawing.Point(0, 0)
        Me.dbcProveedor.Name = "dbcProveedor"
        Me.dbcProveedor.Size = New System.Drawing.Size(121, 21)
        Me.dbcProveedor.TabIndex = 0
        '
        'optGenero
        '
        '
        'optMoneda
        '
        '
        'optMovimiento
        '
        '
        'txtAdicional
        '
        '
        'txtCodigodelProveedor
        '
        '
        'txtCostoAdicional
        '
        '
        'txtCostoFactura
        '
        '
        'txtCostoIndirecto
        '
        '
        'txtCostoReal
        '
        '
        'txtDescripcion
        '
        '
        'txtPrecioenDolares
        '
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Location = New System.Drawing.Point(416, 637)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(93, 35)
        Me.btnLimpiar.TabIndex = 12
        Me.btnLimpiar.Text = "Limpiar"
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.Location = New System.Drawing.Point(309, 637)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(93, 35)
        Me.btnEliminar.TabIndex = 11
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.Location = New System.Drawing.Point(201, 637)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(93, 35)
        Me.btnGuardar.TabIndex = 10
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(522, 637)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(93, 35)
        Me.btnBuscar.TabIndex = 61
        Me.btnBuscar.Text = "Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'frmCorpoABCArticulos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(727, 684)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.txtDescArticulo)
        Me.Controls.Add(Me.txtCodArticulo)
        Me.Controls.Add(Me.sstArticulo)
        Me.Controls.Add(Me._lblArticulo_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(298, 150)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoABCArticulos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Artículos"
        Me.sstArticulo.ResumeLayout(False)
        Me._sstArticulo_TabPage0.ResumeLayout(False)
        Me._fraContenedor_0.ResumeLayout(False)
        Me._fraContenedor_0.PerformLayout()
        Me.fraDiamanteSuelto.ResumeLayout(False)
        Me.fraDiamanteSuelto.PerformLayout()
        Me._fraMoneda_5.ResumeLayout(False)
        Me._fraMoneda_0.ResumeLayout(False)
        Me._fraImagen_0.ResumeLayout(False)
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).EndInit()
        Me._Frame2_0.ResumeLayout(False)
        Me._Frame2_0.PerformLayout()
        Me._Frame4_0.ResumeLayout(False)
        Me._Frame4_0.PerformLayout()
        Me._Frame1_0.ResumeLayout(False)
        Me._Frame1_0.PerformLayout()
        Me._sstArticulo_TabPage1.ResumeLayout(False)
        Me._fraContenedor_1.ResumeLayout(False)
        Me._fraContenedor_1.PerformLayout()
        Me._fraMoneda_3.ResumeLayout(False)
        Me._fraArticulo_1.ResumeLayout(False)
        Me._fraArticulo_2.ResumeLayout(False)
        Me._fraImagen_1.ResumeLayout(False)
        CType(Me.Image2, System.ComponentModel.ISupportInitialize).EndInit()
        Me._fraMoneda_1.ResumeLayout(False)
        Me._Frame1_2.ResumeLayout(False)
        Me._Frame1_2.PerformLayout()
        Me._Frame2_2.ResumeLayout(False)
        Me._Frame2_2.PerformLayout()
        Me._Frame4_1.ResumeLayout(False)
        Me._Frame4_1.PerformLayout()
        Me._sstArticulo_TabPage2.ResumeLayout(False)
        Me._fraContenedor_2.ResumeLayout(False)
        Me._fraContenedor_2.PerformLayout()
        Me._fraMoneda_4.ResumeLayout(False)
        Me._fraMoneda_2.ResumeLayout(False)
        Me._Frame2_3.ResumeLayout(False)
        Me._Frame2_3.PerformLayout()
        Me._Frame4_2.ResumeLayout(False)
        Me._Frame4_2.PerformLayout()
        Me._Frame1_3.ResumeLayout(False)
        Me._Frame1_3.PerformLayout()
        Me._fraImagen_2.ResumeLayout(False)
        CType(Me.Image3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        CType(Me.Frame1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Frame2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Frame4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdBuscarImagen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraArticulo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraContenedor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraImagen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblArticulo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDescripcion, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblMargen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optGenero, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMovimiento, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAdicional, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCodigodelProveedor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCostoAdicional, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCostoFactura, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCostoIndirecto, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCostoReal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDescripcion, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtImagen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPrecioenDolares, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub


End Class