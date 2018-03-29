'**********************************************************************************************************************'
'*PROGRAMA: MENU CORPORATIVO JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA     
'*FECHA DE INICIO: 02/01/2018 
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports System.IO
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class MDIMenuPrincipalCorpo

    Inherits System.Windows.Forms.Form

    Dim r As New Random
    Dim ListImage As New List(Of Image)
    Private components As System.ComponentModel.IContainer

    'Dim PicFolder As String = My.Computer.FileSystem.SpecialDirectories.MyPictures
    'Dim picFolder As String = "C:\Users\Angel Wha\Desktop\CORPORATIVO Y JOYERIA\CODIGO ANGEL WHA\CorporativoV1\CorporativoV1\Resources"
    'Dim dirPath As DirectoryInfo = New DirectoryInfo(picFolder)
    Public ToolTip1 As New System.Windows.Forms.ToolTip
    Public WithEvents Timer1 As System.Windows.Forms.Timer
    Public WithEvents imgEstandar As System.Windows.Forms.ImageList
    Public WithEvents MenuAcercaDe As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents menuContextualGenOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuArchivoOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuBancosOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuBancosOpcCatalogos As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuBancosOpcProcesoDiario As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuBancosOpcProcesoDiarioRptOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuBancosOpcProcesoMensual As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuCatalogosOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuCompyCxPOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuCompyCxPRptOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuConfiguracionOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuConfiguracionOpcUtil As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuContextualOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuEdicionOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuFacturacionOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuFacturacionRptFactOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuInvEntradasOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuInvHojaOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuInvOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuInvRptOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuInvSalidasOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuSegOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuVentasOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuVentasSalMerOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuVentasSalMerOpcRepEjec As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuVentasVendExtOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuVentasVtasIngrOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuVerOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuVerToolBarOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents mnuVerVentanaOpc As New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    Public WithEvents _mnuCatalogosOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_6 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_7 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_8 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_9 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_10 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_11 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_12 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_13 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_14 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_15 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuCatalogosOpc_16 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_17 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_18 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuCatalogosOpc_19 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_20 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_21 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuCatalogosOpc_22 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_23 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCatalogosOpc_24 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuCatalogos As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_6 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_7 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_8 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_9 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_10 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_11 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_12 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuVentasSalMerOpcRepEjec_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpcRepEjec_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpcRepEjec_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpcRepEjec_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpcRepEjec_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpcRepEjec_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasSalMerOpc_13 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasVtasIngrOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasVtasIngrOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasVtasIngrOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasVtasIngrOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasVtasIngrOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasVendExtOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasVendExtOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasVendExtOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasVendExtOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasVendExtOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasVendExtOpc_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasOpc_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasOpc_6 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuVentasOpc_7 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasOpc_8 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasOpc_9 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuVentasOpc_10 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuVentasOpc_11 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuVentas As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFacturacionOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFacturacionOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFacturacionRptFactOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFacturacionRptFactOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFacturacionRptFactOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFacturacionRptFactOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFacturacionRptFactOpc_4 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuFacturacionRptFactOpc_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFacturacionOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFacturacionOpc_3 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuFacturacionOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuFacturacionOpc_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFacturacion As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPOpc_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPOpc_6 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPRptOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPRptOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPRptOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPRptOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPRptOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPRptOpc_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPRptOpc_6 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuCompyCxPRptOpc_7 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPRptOpc_8 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPRptOpc_9 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPRptOpc_10 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPOpc_7 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuCompyCxPOpc_8 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuCompyCxPOpc_9 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuComprasyCxP As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiario_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiario_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiario_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiario_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiario_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiario_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiario_6 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiario_7 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiario_8 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiarioRptOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiarioRptOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiarioRptOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoDiario_9 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoMensual_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoMensual_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoMensual_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoMensual_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoMensual_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcProcesoMensual_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcCatalogos_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcCatalogos_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcCatalogos_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpcCatalogos_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuBancosOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuBancos As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvEntradasOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvEntradasOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvEntradasOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvEntradasOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvEntradasOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvEntradasOpc_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvEntradasOpc_6 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvEntradasOpc_7 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvSalidasOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvSalidasOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvSalidasOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvSalidasOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvSalidasOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvSalidasOpc_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvSalidasOpc_6 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvSalidasOpc_7 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvHojaOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvHojaOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvRptOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvRptOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvRptOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvRptOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvRptOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuInvOpc_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuInventarios As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuConfiguracionOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuConfiguracionOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuConfiguracionOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuConfiguracionOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuConfiguracionOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuConfiguracionOpc_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuConfiguracionOpc_6 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuConfiguracionOpc_7 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuConfiguracionOpc_8 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuConfiguracionOpc_9 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuConfiguracionOpc_10 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuConfiguracionOpc_11 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuConfiguracionOpcUtil_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuConfiguracionOpc_12 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuConfiguracion As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuSegOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuSegOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuSegOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuSeg As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _menuContextualGenOpc_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _menuContextualGenOpc_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _menuContextualGenOpc_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _menuContextualGenOpc_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _menuContextualGenOpc_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _menuContextualGenOpc_5 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _menuContextualGenOpc_6 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _menuContextualGenOpc_7 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _menuContextualGenOpc_8 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _menuContextualGenOpc_9 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _menuContextualGenOpc_10 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents menuContextualGen As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _MenuAcercaDe_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents ButtonContainer As System.Windows.Forms.Panel
    Public WithEvents Panel1 As System.Windows.Forms.Panel
    Public WithEvents lblActualizacion As Label
    Public WithEvents btnSoporte As System.Windows.Forms.Button
    Public WithEvents ButtonTeleMarketing As System.Windows.Forms.Button
    Public WithEvents ButtonCorteDiario As System.Windows.Forms.Button
    Public WithEvents ButtonRegistroCobranza As System.Windows.Forms.Button
    Public WithEvents ButtonRegistroGastos As System.Windows.Forms.Button
    Public WithEvents ButtonConsultaInventario As System.Windows.Forms.Button
    Public WithEvents ButtonCompraEmergencia As System.Windows.Forms.Button
    Public WithEvents ButtonSalidasAOrden As System.Windows.Forms.Button
    Public WithEvents ButtonRecepcionProducto As System.Windows.Forms.Button
    Public WithEvents ButtonOrdenCompra As System.Windows.Forms.Button
    Public WithEvents ButtonHistorial As System.Windows.Forms.Button
    Public WithEvents ButtonCalendario As System.Windows.Forms.Button
    Public WithEvents ButtonCalculadora As System.Windows.Forms.Button
    Public WithEvents ButtonCotizacion As System.Windows.Forms.Button
    Public WithEvents ButtonEmpresas As System.Windows.Forms.Button
    Public WithEvents ButtonClientes As System.Windows.Forms.Button
    Public WithEvents ButtonRecepcion As System.Windows.Forms.Button
    Public WithEvents ButtonPanelControl As System.Windows.Forms.Button
    Public WithEvents panel2 As System.Windows.Forms.Panel
    Public WithEvents lblhora As System.Windows.Forms.Label
    Public WithEvents lblfecha As System.Windows.Forms.Label
    Public WithEvents lbluser As System.Windows.Forms.Label
    Public WithEvents btnmin As System.Windows.Forms.Button
    Public WithEvents btncerrar As System.Windows.Forms.Button
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.

    ''' ***************************************************************************************************************************************************
    ''' MODIFICACION ABC ARTICULOS - MANEJO DIAMANTE SUELTO
    ''' REPORTE VENTAS Y EXISTENCIAS POR FAMILIA
    ''' 27OCT2010 - MAVF Ver
    '''
    ''' Ver 1.0       Estatus : Aprobado
    ''' **************************************************************************************************************************************************
    ''' 
    ' PROGRAMA:  Menu Principal
    ' FECHA   :  08/Mayo/2003

    Dim I As Integer
    Public WithEvents btnmaxi As Button
    Dim AcumuladoTimer As Integer 'Acumula el tiempo que ha transcurrido en el Evento Timer. para que cuando sea 120 min. Ejecute el Evento.

    ' Procedimiento que ejecuta una consulta para dar valor a las variables globales
    ' para Joyería, Relojería y Varios (gCODJOYERIA, gCODRELOJERIA y gVARIOS respectivamente)
    Sub GETCODGRUPOS()
        On Error GoTo Merr
        Dim I As Integer
        gStrSql = "SELECT codGrupo, descGrupo FROM catGrupos"

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'Si no regresa datos la consulta, entonces manda mensage y sale del sistema
        If RsGral.RecordCount > 0 Then
            RsGral.MoveFirst()
            For I = 1 To RsGral.RecordCount
                Select Case True
                    Case Trim(UCase(RsGral.Fields("DescGrupo").Value)) = C_JOYERIA
                        gCODJOYERIA = RsGral.Fields("CodGrupo").Value
                    Case Trim(UCase(RsGral.Fields("DescGrupo").Value)) = C_RELOJERIA
                        gCODRELOJERIA = RsGral.Fields("CodGrupo").Value
                    Case Trim(UCase(RsGral.Fields("DescGrupo").Value)) = C_VARIOS
                        gCODVARIOS = RsGral.Fields("CodGrupo").Value
                End Select
                RsGral.MoveNext()
            Next I
        Else
            MsgBox("No hay Grupos de Artículos debidamente registrados," & vbNewLine & "llame al proveedor del sistema.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error crítico, el programa debe cerrarse ...")
            Me.Close()
        End If
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    '    ' función que indica si hay un archivo de conexion a la bd, si no lo hay
    '    ' manda llamar al formulario de conexion para crearlo
    '    Function Conexion() As Boolean
    '        On Error GoTo Errores
    '        Dim ArchivoTxt, F As Object
    '        Dim miArchivo As String
    '        Dim Server, Bd1 As String

    '        miArchivo = Dir(My.Application.Info.DirectoryPath & "\Sistema\CJoyeria.Txt")
    '        If miArchivo <> "" Then
    '            Conexion = True
    '        Else
    '            MsgBox("El archivo principal de configuración de la conexión no está creado." & vbNewLine & "Se mostrará la ventana de creación del archivo ", MsgBoxStyle.Information, "JOYERIA MARIO RAMOS")
    '            frmVerificarConexion.ShowDialog()
    '            Conexion = True
    '        End If

    '        If Not Conexion Then Exit Function
    '        '''Donde corra la aplicacion - Servidor
    '        miArchivo = Dir(My.Application.Info.DirectoryPath & "\Sistema", FileAttribute.Directory)
    '        If miArchivo = "" Then
    '            MkDir(My.Application.Info.DirectoryPath & "\Sistema")
    '        End If

    '        miArchivo = Dir(My.Application.Info.DirectoryPath & "\Sistema\Imagenes", FileAttribute.Directory)
    '        If miArchivo = "" Then
    '            MkDir(My.Application.Info.DirectoryPath & "\Sistema\Imagenes")
    '        End If

    '        '''Archivo de conexion
    '        ArchivoTxt = CreateObject("Scripting.FileSystemObject")
    '        F = ArchivoTxt.OpenTextFile(My.Application.Info.DirectoryPath & "\Sistema\CJoyeria.Txt", 1, -2)
    '        Dim S, t1 As String
    '        Dim I As Integer
    '        S = F.ReadLine
    '        t1 = F.ReadLine
    '        NombreServidor = S
    '        NombreBaseDatos = t1
    '        F.Close()
    '        If Not ModConexion.Abrir(S, t1) Then
    '            frmVerificarConexion.TxtNomServidor.Text = S
    '            frmVerificarConexion.TxtBDPrincipal.Text = t1
    '            frmVerificarConexion.Show()
    '            Exit Function
    '        Else
    '            Conexion = True
    '        End If

    'Errores:
    '        If Err.Number <> 0 Then
    '            Conexion = False
    '            ModErrores.Errores()
    '        End If
    '    End Function

    '    Public Sub MenuConfiguracionInicial(ByRef Periodo As String)
    '        '''este procedimiento se usa para inabilitar las opciones del
    '        '''menu cuando apenas se esta iniciando sesion y que se va a
    '        '''capturar el login y password de usuario
    '        With Me
    '            ''en esta parte se inabilitan los menus que conforman el
    '            ''sistema corporativo
    '            ''del menu archivo solo esta habilitada la opcion salir
    '            ''por eso ese menu esta habilitado
    '            Select Case Periodo
    '                'Case "I"
    '                '    .mnuArchivoOpc(0).Enabled = False
    '                '    .mnuArchivoOpc(1).Enabled = False
    '                '    .mnuArchivoOpc(2).Enabled = False
    '                '    ''los demas menus estan inabilitados
    '                '    .mnuEdicion.Enabled = False
    '                '    .mnuVer.Enabled = False
    '                '    .mnuCatalogos.Enabled = False
    '                '    .mnuVentas.Enabled = False
    '                '    .mnuFacturacion.Enabled = False

    '                '    .mnuComprasyCxP.Enabled = False

    '                '    .mnuBancos.Enabled = False
    '                '    .mnuInventarios.Enabled = False
    '                '    .mnuConfiguracion.Enabled = False
    '                '    .mnuSeg.Enabled = False
    '                '    With .ToolbarStandar
    '                '        '''en esta parte se inabilitan las opciones de la barra
    '                '        '''de herramientas standar
    '                '        For I = 1 To 7
    '                '            'UPGRADE_WARNING: Lower bound of collection MenuPrincipal.ToolbarStandar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
    '                '            .Items.Item(I).Enabled = False
    '                '        Next I
    '                '    End With
    '                'Case "F"
    '                '    .mnuArchivoOpc(0).Enabled = True
    '                '    .mnuArchivoOpc(1).Enabled = True
    '                '    .mnuArchivoOpc(2).Enabled = True
    '                '    ''los demas menus estan habilitados
    '                '    .mnuEdicion.Enabled = True
    '                '    .mnuVer.Enabled = True
    '                '    .mnuCatalogos.Enabled = True
    '                '    .mnuVentas.Enabled = True
    '                '    .mnuFacturacion.Enabled = True

    '                '    .mnuComprasyCxP.Enabled = True

    '                '    .mnuBancos.Enabled = True
    '                '    .mnuInventarios.Enabled = True
    '                '    .mnuConfiguracion.Enabled = True
    '                '    .mnuSeg.Enabled = True
    '                '    With .ToolbarStandar
    '                '        '''en esta parte se habilitan las opciones de la barra
    '                '        '''de herramientas standar
    '                '        For I = 1 To 7
    '                '            'UPGRADE_WARNING: Lower bound of collection MenuPrincipal.ToolbarStandar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
    '                '            .Items.Item(I).Enabled = True
    '                '        Next I
    '                '    End With
    '            End Select
    '        End With
    '    End Sub

    '    Private Sub MDIForm_Initialize_Renamed()
    '        'UPGRADE_NOTE: Object frmBancosProcesoDiarioOrigenyAplicacion may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '        'frmBancosProcesoDiarioOrigenyAplicacion = Nothing
    '        'frmBancosProcesoDiarioOrigenyAplicacion.Close()
    '    End Sub



    '    Private Sub MenuPrincipal_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    '        'If frmFactAnalisisVentas.gblnCambiosAnalisis Then
    '        '    MsgBox("No es posible limpiar la pantalla, ya ha generado algun(s) folios adicionales" & vbNewLine & "  Para poder limpiar la pantalla debera generar la factura correspondiente", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
    '        '    Cancel = 1
    '        '    Exit Sub
    '        'End If
    '        ModConexion.Cerrar()
    '        'UPGRADE_NOTE: Object MenuPrincipal may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '        'Me = Nothing
    '        End
    '    End Sub

    Public Sub MenuAcercaDe_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MenuAcercaDe.Click
        Dim Index As Integer = MenuAcercaDe.GetIndex(eventSender)
        Me.IsMdiContainer = True
        frmAcercaDe.MdiParent = Me
        frmAcercaDe.Show()
    End Sub

    Public Sub mnuArchivoOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuArchivoOpc.Click
        Dim Index As Integer = mnuArchivoOpc.GetIndex(eventSender)
        On Error Resume Next
        Select Case Index
            Case 0 'Guardar
                'ActiveMdiChild.Guardar()
            Case 1 'Imprimir
                'ActiveMdiChild.Imprime()
            Case 2 'Cerrar
                'System.Windows.Forms.Form.ActiveForm.Close()
                'Me.Close()
                'Dim frmAcceso As FrmAcceso = New FrmAcceso()
                'frmAcceso.Show()
            Case 4 'Salir
                'If frmFactAnalisisVentas.gblnCambiosAnalisis Then
                '    MsgBox("No es posible limpiar la pantalla, ya ha generado algun(s) folios adicionales" & vbNewLine & "  Para poder limpiar la pantalla debera generar la factura correspondiente", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                '    Exit Sub
                'End If
                End
        End Select
    End Sub

    Public Sub mnuBancosOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuBancosOpc.Click
        Dim Index As Integer = mnuBancosOpc.GetIndex(eventSender)
        Select Case Index
            Case 2 'Depuración de Movimientos Historicos
                Me.IsMdiContainer = True
                frmBancosProcesoMensualDepuraciondeMovimientosHist.MdiParent = Me
                frmBancosProcesoMensualDepuraciondeMovimientosHist.Show()
                'frmBancosProcesoMensualDepuraciondeMovimientosHist.BringToFront()
        End Select
    End Sub

    Public Sub mnuBancosOpcCatalogos_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuBancosOpcCatalogos.Click
        Dim Index As Integer = mnuBancosOpcCatalogos.GetIndex(eventSender)
        Dim I As Integer
        Dim Found As Boolean

        Select Case Index
            Case 0 ' ABC  a Bancos
                Found = False
                For I = 0 To My.Application.OpenForms.Count - 1
                    If UCase(Trim(My.Application.OpenForms.Item(I).Name)) = "FRMCORPOABCBANCOS" Then
                        Found = True
                    End If
                Next
                If Not Found Then
                    FrmAbcBancos.Show()
                    FrmAbcBancos.BringToFront()
                End If

            Case 1 ' ABC  a Cuentas Bancarias
                Found = False
                For I = 0 To My.Application.OpenForms.Count - 1
                    If UCase(Trim(My.Application.OpenForms.Item(I).Name)) = "FRMCORPOABCCUENTASBANCARIAS" Then
                        Found = True
                    End If
                Next
                If Not Found Then
                    frmAbcCuentasBancarias.Show()
                    frmAbcCuentasBancarias.BringToFront()
                End If

            Case 2 'ABC a Origen y Aplicación de Recursos
                Found = False
                For I = 0 To My.Application.OpenForms.Count - 1
                    If UCase(Trim(My.Application.OpenForms.Item(I).Name)) = "FRMCORPOABCORIGENYAPLICACIONDERECURSOS" Then
                        Found = True
                    End If
                Next
                If Not Found Then
                    frmAbcOrigenyAplicaciondeRecursos.Show()
                    frmAbcOrigenyAplicaciondeRecursos.BringToFront()
                End If

            Case 3 'ABC a Rubros de Origen y Aplicación
                Found = False
                For I = 0 To My.Application.OpenForms.Count - 1
                    If UCase(Trim(My.Application.OpenForms.Item(I).Name)) = "FRMCORPOABCRUBROSDEAPLICACIONYORIGEN" Then
                        Found = True
                    End If
                Next
                If Not Found Then
                    Me.IsMdiContainer = True
                    frmCorpoABCRubrosdeAplicacionyOrigen.MdiParent = Me
                    frmCorpoABCRubrosdeAplicacionyOrigen.Show()
                    'frmAbcRubrosdeAplicacionyOrigen.Show()
                    'frmAbcRubrosdeAplicacionyOrigen.BringToFront()
                End If

        End Select


    End Sub

    Public Sub mnuBancosOpcProcesoDiario_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuBancosOpcProcesoDiario.Click
        Dim Index As Integer = mnuBancosOpcProcesoDiario.GetIndex(eventSender)
        Select Case Index
            Case 0 'Registro de Pagos
                Me.IsMdiContainer = True
                frmBancosProcesoDiarioRegistrodePagos.MdiParent = Me
                frmBancosProcesoDiarioRegistrodePagos.Show()
            Case 1 'Registro de Depositos
                Me.IsMdiContainer = True
                frmBancosProcesoDiarioRegistrodeDepositos.MdiParent = Me
                frmBancosProcesoDiarioRegistrodeDepositos.Show()
                'frmBancosProcesoDiarioRegistrodeDepositos.BringToFront()
            Case 2 'Registro de Cargos Diversos
                Me.IsMdiContainer = True
                frmBancosProcesoDiarioCargosDiversos.MdiParent = Me
                frmBancosProcesoDiarioCargosDiversos.Show()
                'frmBancosProcesoDiarioCargosDiversos.BringToFront()
            Case 3 'Registro de Traspasos Bancarios
                Me.IsMdiContainer = True
                frmBancosProcesoDiarioTraspasosBancarios.MdiParent = Me
                frmBancosProcesoDiarioTraspasosBancarios.Show()
                'frmBancosProcesoDiarioTraspasosBancarios.BringToFront()
            Case 4 'Registro de Anticipo a Proveedores/Acreedores
                Me.IsMdiContainer = True
                frmBancosProcesoDiarioAnticipoProveedoresAcreed.MdiParent = Me
                frmBancosProcesoDiarioAnticipoProveedoresAcreed.Show()
                'frmBancosProcesoDiarioAnticipoProveedoresAcreed.BringToFront()
            Case 5 'Registro de Otros Ingresos
                Me.IsMdiContainer = True
                frmBancosProcesoDiarioRegistrodeOtrosIngresos.MdiParent = Me
                frmBancosProcesoDiarioRegistrodeOtrosIngresos.Show()
                'frmBancosProcesoDiarioRegistrodeOtrosIngresos.BringToFront()
            Case 6 'Cancelar Movimientos
                Me.IsMdiContainer = True
                frmBancosProcesoDiarioCancelaciondeMovimientosBanc.MdiParent = Me
                frmBancosProcesoDiarioCancelaciondeMovimientosBanc.Show()
                'frmBancosProcesoDiarioCancelaciondeMovimientosBanc.BringToFront()
            Case 7 'Consulta de Saldos
                Me.IsMdiContainer = True
                frmBancosProcesoDiarioConsultadeSaldos.MdiParent = Me
                frmBancosProcesoDiarioConsultadeSaldos.Show()
                'frmBancosProcesoDiarioConsultadeSaldos.BringToFront()
            Case 8 'Cierre Diario de Bancos
                Me.IsMdiContainer = True
                frmBancosProcesoDiarioCierreDiarioBancos.MdiParent = Me
                frmBancosProcesoDiarioCierreDiarioBancos.Show()
        End Select
    End Sub

    Public Sub mnuBancosOpcProcesoDiarioRptOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuBancosOpcProcesoDiarioRptOpc.Click
        Dim Index As Integer = mnuBancosOpcProcesoDiarioRptOpc.GetIndex(eventSender)
        Select Case Index
            Case 0 'Rpt Movtos Banc 
                Me.IsMdiContainer = True
                frmBancosReportedeMovimientosBancarios.MdiParent = Me
                frmBancosReportedeMovimientosBancarios.Show()
            Case 1 'Rpt Movtos Banc x Tipo 
                Me.IsMdiContainer = True
                frmBancosReporteMovBancariosXTipo.MdiParent = Me
                frmBancosReporteMovBancariosXTipo.Show()
            Case 2 'Analisis Diario de Bancos 
                Me.IsMdiContainer = True
                frmBancosReporteAnalisisDiarioBancos.MdiParent = Me
                frmBancosReporteAnalisisDiarioBancos.Show()
        End Select
    End Sub

    Public Sub mnuBancosOpcProcesoMensual_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuBancosOpcProcesoMensual.Click
        Dim Index As Integer = mnuBancosOpcProcesoMensual.GetIndex(eventSender)
        Select Case Index
            Case 0 'Conciliación Manual
                Me.IsMdiContainer = True
                frmBancosProcesoMensualConciliacionMensual.MdiParent = Me
                frmBancosProcesoMensualConciliacionMensual.Show()
                frmBancosProcesoMensualConciliacionMensual.BringToFront()
            Case 1 'Reporte de Movimientos en Conciliación
                Me.IsMdiContainer = True
                frmBancosProcesoMensualMovimientosenConciliacion.MdiParent = Me
                frmBancosProcesoMensualMovimientosenConciliacion.Show()
                frmBancosProcesoMensualMovimientosenConciliacion.BringToFront()
            Case 2 'Flujo de Caja General
                Me.IsMdiContainer = True
                frmBancosProcesoMensualFlujoCajaGeneral.MdiParent = Me
                frmBancosProcesoMensualFlujoCajaGeneral.Show()
                frmBancosProcesoMensualFlujoCajaGeneral.BringToFront()
            Case 3 'Consulta de Origen y Aplicación
                Me.IsMdiContainer = True
                frmBancosProcesoMensualConsultaOrigenAplicRec.MdiParent = Me
                frmBancosProcesoMensualConsultaOrigenAplicRec.Show()
                frmBancosProcesoMensualConsultaOrigenAplicRec.BringToFront()
            Case 4 'Reporte de Origen y Aplicación
                Me.IsMdiContainer = True
                frmBancosProcesoMensualReporteOrigenyAplicacion.MdiParent = Me
                frmBancosProcesoMensualReporteOrigenyAplicacion.Show()
                frmBancosProcesoMensualReporteOrigenyAplicacion.BringToFront()
            Case 5
                Me.IsMdiContainer = True
                frmBancosProcesoMensualCierreConciliacion.MdiParent = Me
                frmBancosProcesoMensualCierreConciliacion.Show()
                frmBancosProcesoMensualCierreConciliacion.BringToFront()
        End Select
    End Sub

    Public Sub mnuCatalogosOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCatalogosOpc.Click
        Dim Index As Integer = mnuCatalogosOpc.GetIndex(eventSender)
        Dim I As Integer
        Dim Found As Boolean

        Select Case Index
            Case 0 'ABC de Clientes
                Me.IsMdiContainer = True
                frmCorpoABCClientes.MdiParent = Me
                frmCorpoABCClientes.Show()
            Case 1 'ABC de Vendedores 
                Me.IsMdiContainer = True
                FrmCorpoAbcVendedores.MdiParent = Me
                FrmCorpoAbcVendedores.Show()
            Case 2 'ABC a Sucursales (Almacenes de forma interna)
                Me.IsMdiContainer = True
                frmCorpoAbcSucursales.MdiParent = Me
                frmCorpoAbcSucursales.Show()
            Case 3 ' ABC a Talleres 
                Me.IsMdiContainer = True
                frmCorpoAbcTalleres.MdiParent = Me
                frmCorpoAbcTalleres.Show()
            Case 4 ' ABC a Tipos de Material 
                Me.IsMdiContainer = True
                frmCorpoAbcTiposMaterial.MdiParent = Me
                frmCorpoAbcTiposMaterial.Show()
            Case 5 ' ABC a Prov y Acreed 
                Me.IsMdiContainer = True
                frmCorpoAbcProvAcreed.MdiParent = Me
                frmCorpoAbcProvAcreed.Show()
            Case 6 ' ABC  a formas de pago 
                Me.IsMdiContainer = True
                frmCorpoAbcFormasdePago.MdiParent = Me
                frmCorpoAbcFormasdePago.Show()
            Case 7 ' ABC  a Bancos
                Found = False
                For I = 0 To My.Application.OpenForms.Count - 1
                    If UCase(Trim(My.Application.OpenForms.Item(I).Name)) = "FRMCORPOABCBANCOS" Then
                        Found = True
                    End If
                Next
                If Not Found Then
                    Me.IsMdiContainer = True
                    FrmAbcBancos.MdiParent = Me
                    FrmAbcBancos.Show()
                End If

            Case 8 ' ABC  a Cuentas Bancarias
                Found = False
                For I = 0 To My.Application.OpenForms.Count - 1
                    If UCase(Trim(My.Application.OpenForms.Item(I).Name)) = "FRMCORPOABCCUENTASBANCARIAS" Then
                        Found = True
                    End If
                Next
                If Not Found Then
                    Me.IsMdiContainer = True
                    frmAbcCuentasBancarias.MdiParent = Me
                    frmAbcCuentasBancarias.Show()
                End If

            Case 9 'ABC a Origen y Aplicación de Recursos
                Found = False
                For I = 0 To My.Application.OpenForms.Count - 1
                    If UCase(Trim(My.Application.OpenForms.Item(I).Name)) = "FRMCORPOABCORIGENYAPLICACIONDERECURSOS" Then
                        Found = True
                    End If
                Next
                If Not Found Then
                    Me.IsMdiContainer = True
                    frmAbcOrigenyAplicaciondeRecursos.MdiParent = Me
                    frmAbcOrigenyAplicaciondeRecursos.Show()
                End If

            Case 10 'ABC a Rubros de Origen y Aplicación
                Found = False
                For I = 0 To My.Application.OpenForms.Count - 1
                    If UCase(Trim(My.Application.OpenForms.Item(I).Name)) = "FRMCORPOABCRUBROSDEAPLICACIONYORIGEN" Then
                        Found = True
                    End If
                Next
                If Not Found Then
                    Me.IsMdiContainer = True
                    frmCorpoABCRubrosdeAplicacionyOrigen.MdiParent = Me
                    frmCorpoABCRubrosdeAplicacionyOrigen.Show()
                End If
            Case 11 'ABC a Descuentos a Vendedores Externos 
                Me.IsMdiContainer = True
                frmCorpoABCDescuentosVendExternos.MdiParent = Me
                frmCorpoABCDescuentosVendExternos.Show()
            Case 12 'ABC a Promociones de Tarjetas 
                Me.IsMdiContainer = True
                frmCorpoAbcPromocionesTarjetasBanc.MdiParent = Me
                frmCorpoAbcPromocionesTarjetasBanc.Show()
            Case 13 'ABC a Programacion de ofertas 
                Me.IsMdiContainer = True
                frmProgramacionPromociones.MdiParent = Me
                frmProgramacionPromociones.Show()
            Case 14 'ABC de comisiones por ventas de vendedores 
                Me.IsMdiContainer = True
                frmCorpoABCComisiones.MdiParent = Me
                frmCorpoABCComisiones.Show()
            Case 16 'ABC a Artículos 
                Me.IsMdiContainer = True
                frmCorpoABCArticulos.MdiParent = Me
                frmCorpoABCArticulos.Show()
            Case 17 'CC a Grupos de Artículos 
                Me.IsMdiContainer = True
                frmCorpoABCGrupos.MdiParent = Me
                frmCorpoABCGrupos.Show()
            Case 19 'ABC a Marcas de Relojería 
                Me.IsMdiContainer = True
                frmCorpoABCMarca.MdiParent = Me
                frmCorpoABCMarca.Show()
            Case 20 'ABC a Modelos de Relojería 
                Me.IsMdiContainer = True
                frmCorpoABCModelos.MdiParent = Me
                frmCorpoABCModelos.Show()
            Case 22 'ABC a Familias de Artículos 
                Me.IsMdiContainer = True
                frmCorpoABCFamilias.MdiParent = Me
                frmCorpoABCFamilias.Show()
            Case 23 'ABC a Líneas de Artículos
                Me.IsMdiContainer = True
                frmCorpoABCLineas.MdiParent = Me
                frmCorpoABCLineas.Show()
            Case 24 'ABC a SubLíneas de Joyería
                Me.IsMdiContainer = True
                frmCorpoABCSubLineas.MdiParent = Me
                frmCorpoABCSubLineas.Show()
        End Select
    End Sub

    Public Sub mnuCompyCxPOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCompyCxPOpc.Click
        Dim Index As Integer = mnuCompyCxPOpc.GetIndex(eventSender)
        'Menu: Compras y CxP
        Select Case Index
            Case 0 'Orden de Compra
                frmCXPOrdenCompra.Show()
                frmCXPOrdenCompra.BringToFront()
            Case 1 'Registro de Facturas de Compras
                'frmCXPRegFactCompras.Show()
                'frmCXPRegFactCompras.BringToFront()
            Case 2 'registro de Facturas de Gastos
                'frmCXPRegFactGastos.Show()
                'frmCXPRegFactGastos.BringToFront()
            Case 3 'Programación Especial de Pagos
                'frmCXPProgPagos.Show()
                'frmCXPProgPagos.BringToFront()
            Case 4 'Notas de Crédito
                'frmCXPRegNotasCredito.Show()
                'frmCXPRegNotasCredito.BringToFront()
            Case 5 'Cuentas por Pagar
                'frmCXPCuentasPorPagar.Show
                'frmCXPCuentasPorPagar.ZOrder
                frmCXPReporteCuentasporPagar.Show()
                frmCXPReporteCuentasporPagar.BringToFront()
            Case 6 'Emisión de Pagos
                frmCXPEmisionPagos.Show()
                frmCXPEmisionPagos.BringToFront()
            Case 7 'Reportes
            Case 9 '''Carga Inicial de Facturas
                frmCXPRegFactComprasCargaInicial.Show()
                frmCXPRegFactComprasCargaInicial.BringToFront()
        End Select
    End Sub

    Public Sub mnuCompyCxPRptOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCompyCxPRptOpc.Click
        Dim Index As Integer = mnuCompyCxPRptOpc.GetIndex(eventSender)
        'Menu: Compras y CxP - Reportes
        Select Case Index
            Case 0 'Reportes de Facturas
                frmCXPrptFacturas.Show()
                frmCXPrptFacturas.BringToFront()
            Case 1 'Reporte de Notas de Crédito
                frmCXPrptNotasCredito.Show()
                frmCXPrptNotasCredito.BringToFront()
            Case 2 'Reporte de los Mejores
                frmCXPrptMejoresProv.Show()
                frmCXPrptMejoresProv.BringToFront()
            Case 3 'Análisis Anual de Compras
                frmCXPrptComprasPorProveedor.Show()
                frmCXPrptComprasPorProveedor.BringToFront()
            Case 4 'Órdenes de Compra
                frmCXPrptOC.Show()
                frmCXPrptOC.BringToFront()
            Case 5 'Artículos Pendientes por Recibir
                frmCXPrptArticulosPendientes.Show()
                frmCXPrptArticulosPendientes.BringToFront()
            Case 7 'Cuentas por Pagar
                frmCXPReporteCuentasporPagar.Show()
                frmCXPReporteCuentasporPagar.BringToFront()
            Case 8 'Auxiliar de Proveedores
                frmCxpReporteSaldoXProveedor.Show()
                frmCxpReporteSaldoXProveedor.BringToFront()
            Case 9 'CxP Presupuestado
                frmCXPPresupuestado.Show()
                frmCXPPresupuestado.BringToFront()
            Case 10 'Saldos de Proveedores
                frmCXPReporteSaldoXProveedores.Show()
                frmCXPReporteSaldoXProveedores.BringToFront()
        End Select
    End Sub


    Public Sub mnuConfiguracionOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuConfiguracionOpc.Click
        Dim Index As Integer = mnuConfiguracionOpc.GetIndex(eventSender)
        Select Case Index
            Case 0
                Me.IsMdiContainer = True
                frmConfigGralCorporativo.MdiParent = Me
                frmConfigGralCorporativo.Show()
            Case 1
                'Me.IsMdiContainer = True
                'frmPVConfigPuntoVenta.MdiParent = Me
                frmPVConfigPuntoVenta.Show()
            Case 2
                'Me.IsMdiContainer = True
                'frmPVConfigTicketVenta.MdiParent = Me
                'frmPVConfigTicketVenta.Show()
            Case 3
                'Me.IsMdiContainer = True
                'frmPVConfigFacturacion.ShowDialog()
                'frmPVConfigFacturacion.Show()
            Case 4
                Me.IsMdiContainer = True
                frmPVConfigFolios.MdiParent = Me
                frmPVConfigFolios.Show()
            Case 5
                Me.IsMdiContainer = True
                frmPVConfigCaja.MdiParent = Me
                frmPVConfigCaja.Show()
            Case 7
                'Me.IsMdiContainer = True
                'frmConfiguracion.MdiParent = Me
                'frmConfiguracion.Show()
                'Dim frmConfiguracion As New frmConfiguracion()
                frmConfiguracion.Show()
            Case 9
                'Cambiar Usuario 
                Me.Hide()
                'Me.IsMdiContainer = True
                'FrmAcceso.MdiParent = Me
                FrmAcceso.Show()
            Case 10
                'Cambiar contraseña 
                Me.IsMdiContainer = True
                frmCambioPassword.MdiParent = Me
                frmCambioPassword.Show()
        End Select
    End Sub

    Public Sub mnuConfiguracionOpcUtil_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuConfiguracionOpcUtil.Click
        Dim Index As Integer = mnuConfiguracionOpcUtil.GetIndex(eventSender)
        Select Case Index
            Case 0
                Me.IsMdiContainer = True
                frmImportacionImagenes.MdiParent = Me
                frmImportacionImagenes.Show()
        End Select
    End Sub

    '    Public Sub mnuEdicionOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEdicionOpc.Click
    '        Dim Index As Integer = mnuEdicionOpc.GetIndex(eventSender)
    '        On Error Resume Next
    '        Select Case Index
    '            Case 0 'Nuevo
    '                 ActiveMdiChild.Limpiar()
    '            Case 1 'Cancelar
    '                 ActiveMdiChild.Cancelar()
    '            Case 2 'Eliminar
    '                  ActiveMdiChild.Eliminar()
    '            Case 4 'Buscar
    '                 ActiveMdiChild.Buscar()
    '        End Select
    '    End Sub

    Public Sub mnuFacturacionOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFacturacionOpc.Click
        Dim Index As Integer = mnuFacturacionOpc.GetIndex(eventSender)
        'Menu: Facturacion
        Select Case Index
            Case 0 'Análisis de las ventas 
                Me.IsMdiContainer = True
                frmFactAnalisisVentas.MdiParent = Me
                frmFactAnalisisVentas.Show()
            Case 1 'Facturación Especial 
                Me.IsMdiContainer = True
                frmFactFacturacionEspecial.MdiParent = Me
                frmFactFacturacionEspecial.Show()
            Case 4 'Reimpresion de cortes 
                Me.IsMdiContainer = True
                frmReimpresionCorteFinal.MdiParent = Me
                frmReimpresionCorteFinal.Show()
            Case 5 'Diario de movtos 
                Me.IsMdiContainer = True
                frmPVDiarioMovtos.MdiParent = Me
                frmPVDiarioMovtos.Show()
        End Select
    End Sub

    Public Sub mnuFacturacionRptFactOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFacturacionRptFactOpc.Click
        Dim Index As Integer = mnuFacturacionRptFactOpc.GetIndex(eventSender)
        'Menu: Facturacion - Reportes de Facturación
        Select Case Index
            Case 0 'Global por Tienda 
                Me.IsMdiContainer = True
                frmFactReportesFacturacionGlobalXSucursal.MdiParent = Me
                frmFactReportesFacturacionGlobalXSucursal.Show()
            Case 1 'Detallada por tienda 
                Me.IsMdiContainer = True
                frmFactReportesFacturacionDetalladaXSucursal.MdiParent = Me
                frmFactReportesFacturacionDetalladaXSucursal.Show()
            Case 2 'Reimpresión de Tickets 
                Me.IsMdiContainer = True
                frmFactReportesImpresionTickets.MdiParent = Me
                frmFactReportesImpresionTickets.Show()
            Case 3 'Los Mejores Clientes 
                Me.IsMdiContainer = True
                frmFactReportesMejoresClientes.MdiParent = Me
                frmFactReportesMejoresClientes.Show()
            Case 5
                Me.IsMdiContainer = True
                frmFactReporteVtasTarjetaCreditoXSucursal.MdiParent = Me
                frmFactReporteVtasTarjetaCreditoXSucursal.Show()
        End Select
    End Sub

    Public Sub mnuInvEntradasOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuInvEntradasOpc.Click
        Dim Index As Integer = mnuInvEntradasOpc.GetIndex(eventSender)
        'Menu: Inventarios - Entradas
        Select Case Index
            Case 0 'Por Compra 
                'Me.IsMdiContainer = True
                'frmInventradaPorCompra.MdiParent = Me
                'frmInventradaPorCompra.Show()
            Case 1 'Transferencia 
                'Me.IsMdiContainer = True
                'frmInvEntradaPorTransferencia.MdiParent = Me
                'frmInvEntradaPorTransferencia.Show()
            Case 2 'PorDevolución sobre Venta 
                'Me.IsMdiContainer = True
                'frmInvEntradaPorDevolSobreVenta.MdiParent = Me
                'frmInvEntradaPorDevolSobreVenta.Show()
            Case 3 'Por DEvolución de Vendedores Externos 
                'Me.IsMdiContainer = True
                'frmInvEntradaPorDevoldeVendedoresExternos.MdiParent = Me
                'frmInvEntradaPorDevoldeVendedoresExternos.Show()
            Case 4 'Por devolución sobre venta a Consignación 
                'Me.IsMdiContainer = True
                'frmInvEntradaPorDevolsobreVentaAVendedoresExternos.MdiParent = Me
                'frmInvEntradaPorDevolsobreVentaAVendedoresExternos.Show()
            Case 5 'Devolución sobre Obsequio 
                'Me.IsMdiContainer = True
                'frmInvEntradaPorDevolSobreObsequio.MdiParent = Me
                'frmInvEntradaPorDevolSobreObsequio.Show()
            Case 6 'Deevolución de Préstamo 
                'Me.IsMdiContainer = True
                'frmInvEntradaPorDevolPorPrestamo.MdiParent = Me
                'frmInvEntradaPorDevolPorPrestamo.Show()
            Case 7 'Por Ajuste 
                'Me.IsMdiContainer = True
                'frmInvEntradaporAjuste.MdiParent = Me
                'frmInvEntradaporAjuste.Show()
        End Select
    End Sub

    Public Sub mnuInvHojaOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuInvHojaOpc.Click
        Dim Index As Integer = mnuInvHojaOpc.GetIndex(eventSender)
        'Menu: Inventarios
        Select Case Index
            Case 0
                Me.IsMdiContainer = True
                frminvHojadecontrol.MdiParent = Me
                frminvHojadecontrol.Show()
            Case 1
                Me.IsMdiContainer = True
                frmInvAnalisisComparativo.MdiParent = Me
                frmInvAnalisisComparativo.Show()
        End Select
    End Sub

    Public Sub mnuInvOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuInvOpc.Click
        Dim Index As Integer = mnuInvOpc.GetIndex(eventSender)
        'Menu: Inventarios
        Select Case Index
                '        Case 0  'Entradas
                '        Case 1  'salidas
            Case 2 'Impresión de Etiquetas 
                Me.IsMdiContainer = True
                frmImpresionEtiquetas.MdiParent = Me
                frmImpresionEtiquetas.Show()
            Case 3 'Stock Básico de Tienda 
                Me.IsMdiContainer = True
                frmStockBasicoTienda.MdiParent = Me
                frmStockBasicoTienda.Show()
                '        Case 4  'Reportes
        End Select
    End Sub

    Public Sub mnuInvRptOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuInvRptOpc.Click
        Dim Index As Integer = mnuInvRptOpc.GetIndex(eventSender)
        'Menu: Inventarios - Reportes
        Select Case Index
            Case 0 'Kardex 
                Me.IsMdiContainer = True
                frmRptKardexArticulo.MdiParent = Me
                frmRptKardexArticulo.Show()
            Case 1 'Existencias y Costos 
                Me.IsMdiContainer = True
                frmRptExistenciasyCostos.MdiParent = Me
                frmRptExistenciasyCostos.Show()
            Case 2 'Préstamos Pendientes 
                Me.IsMdiContainer = True
                frmReportePrestamosPendientes.MdiParent = Me
                frmReportePrestamosPendientes.Show()
            Case 3 'Compracion Existencia Stock 
                Me.IsMdiContainer = True
                frmrptComparacionExistenciaStock.MdiParent = Me
                frmrptComparacionExistenciaStock.Show()
            Case 4 'Transferencias no conciliadas 
                Me.IsMdiContainer = True
                frmRptTransferenciasNoConciliadas.MdiParent = Me
                frmRptTransferenciasNoConciliadas.Show()
        End Select
    End Sub

    Public Sub mnuInvSalidasOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuInvSalidasOpc.Click
        Dim Index As Integer = mnuInvSalidasOpc.GetIndex(eventSender)
        'Menu: Inventarios - Salidas
        Select Case Index
            Case 0 'Por Venta 
                   'Me.IsMdiContainer = True
                'frmInvSalidaPorVenta.MdiParent = Me
                'frmInvSalidaPorVenta.Show()
            Case 1 'Transferencia 
                  'Me.IsMdiContainer = True
                'frmInvSalidaPorTransferencia.MdiParent = Me
                'frmInvSalidaPorTransferencia.Show()
            Case 2 'Por Devolución sobre Compra 
                       'Me.IsMdiContainer = True
                'frmInvSalidaPorDevolSobreCompra.MdiParent = Me
                'frmInvSalidaPorDevolSobreCompra.Show()
            Case 3 'Salida a Vendedores Externos 
                       'Me.IsMdiContainer = True
                'frmInvSalidaAVendedoresExternos.MdiParent = Me
                'frmInvSalidaAVendedoresExternos.Show()
            Case 4 'Por Venta de VEndedores Externos 
                       'Me.IsMdiContainer = True
                'frmInvSalidaPorVentaAVendedoresExternos.MdiParent = Me
                'frmInvSalidaPorVentaAVendedoresExternos.Show()
            Case 5 'Por Obsequio 
                       'Me.IsMdiContainer = True
                'frmInvSalidaPorObsequio.MdiParent = Me
                'frmInvSalidaPorObsequio.Show()
            Case 6 'Préstamo de Artículo 
                       'Me.IsMdiContainer = True
                'frmInvSalidaPorPrestamo.MdiParent = Me
                'frmInvSalidaPorPrestamo.Show()
            Case 7 'Por Ajuste 
                'Me.IsMdiContainer = True
                'frmInvSalidaporAjuste.MdiParent = Me
                'frmInvSalidaporAjuste.Show()
        End Select
    End Sub

    Public Sub mnuSegOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSegOpc.Click
        Dim Index As Integer = mnuSegOpc.GetIndex(eventSender)
        Select Case Index
            Case 0
                Me.IsMdiContainer = True
                frmRastreoFunciones.MdiParent = Me
                frmRastreoFunciones.Show()
            Case 1
                Me.IsMdiContainer = True
                frmABCModulos.MdiParent = Me
                frmABCModulos.Show()
            Case 2
                Me.IsMdiContainer = True
                frmABCUsuarios.MdiParent = Me
                frmABCUsuarios.Show()
        End Select
    End Sub

    Public Sub mnuVentasOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVentasOpc.Click
        Dim Index As Integer = mnuVentasOpc.GetIndex(eventSender)
        'Menu: Ventas
        Select Case Index
                '        Case 0  'Ventas Salida de Mercancia
                '        Case 1  'Ventas Ingresos
                '        Case 2  'Vendedor Externo
            Case 3 'Apartados 
                Me.IsMdiContainer = True
                frmVtasReportedeApartados.MdiParent = Me
                frmVtasReportedeApartados.Show()
            Case 4 'Reparaciones 
                Me.IsMdiContainer = True
                frmVtasRptReparaciones_Corpo.MdiParent = Me
                frmVtasRptReparaciones_Corpo.Show()
            Case 5 'Cuentas por Cobrar 
                Me.IsMdiContainer = True
                frmVtasReportedeCuentasporCobrar.MdiParent = Me
                frmVtasReportedeCuentasporCobrar.Show()
            Case 7 'Estado de Resultados 
                Me.IsMdiContainer = True
                frmVtasEstadodeResultados.MdiParent = Me
                frmVtasEstadodeResultados.Show()
            Case 8 'Relacion de Gastos 
                Me.IsMdiContainer = True
                frmVtasRelacionGastos.MdiParent = Me
                frmVtasRelacionGastos.Show()
            Case 10 'Control de Reparaciones 
                Me.IsMdiContainer = True
                frmCorpoControlReparaciones_Corpo.MdiParent = Me
                frmCorpoControlReparaciones_Corpo.Show()
            Case 11 'Verificador de Precios 
                Me.IsMdiContainer = True
                frmVerificadorPrecios.MdiParent = Me
                frmVerificadorPrecios.Show()
        End Select
    End Sub

    Public Sub mnuVentasSalMerOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVentasSalMerOpc.Click
        Dim Index As Integer = mnuVentasSalMerOpc.GetIndex(eventSender)
        'Menu: Ventas - Ventas Salida de Mercancia
        Select Case Index
            Case 0 'Por Periodo y Tienda 
                Me.IsMdiContainer = True
                frmVtasRPTVentasSalidadeMercancia.MdiParent = Me
                frmVtasRPTVentasSalidadeMercancia.Show()
            Case 1 'Por Proveedor y Tienda 
                Me.IsMdiContainer = True
                frmVtasRPTVentasSalidadeMercanciaPorProv.MdiParent = Me
                frmVtasRPTVentasSalidadeMercanciaPorProv.Show()
            Case 2 'Por Clasificacion 
                Me.IsMdiContainer = True
                frmVtasRPTVentasSalidadeMercanciaClasifArtic.MdiParent = Me
                frmVtasRPTVentasSalidadeMercanciaClasifArtic.Show()
            Case 3 'COmparativo de Ventas Diarias con año anterior 
                Me.IsMdiContainer = True
                frmVtasRPTVentasSalidadeMercanciaCompara.MdiParent = Me
                frmVtasRPTVentasSalidadeMercanciaCompara.Show()
            Case 4 'Utilidad por Línea 
                Me.IsMdiContainer = True
                frmVtasRPTVentasSalidadeMercanciaUtilidad.MdiParent = Me
                frmVtasRPTVentasSalidadeMercanciaUtilidad.Show()
            Case 5 'Relojería por Marca y Modelos 
                Me.IsMdiContainer = True
                frmVtasRPTVentasSalidadeMercanciaRelojeria.MdiParent = Me
                frmVtasRPTVentasSalidadeMercanciaRelojeria.Show()
            Case 6 'Relojería por Material de Fabricación 
                Me.IsMdiContainer = True
                frmVtasRPTVentasSalidadeMercanciaRelojMaterial.MdiParent = Me
                frmVtasRPTVentasSalidadeMercanciaRelojMaterial.Show()
            Case 7 'Flujo de Venta por Proveedor 
                Me.IsMdiContainer = True
                frmVtasRPTVentasSalidadeMercanciaFlujoVenta.MdiParent = Me
                frmVtasRPTVentasSalidadeMercanciaFlujoVenta.Show()
            Case 8 'Ventas por Cliente 
                Me.IsMdiContainer = True
                frmVtasRPTVentasSalidadeMercanciaPorCliente.MdiParent = Me
                frmVtasRPTVentasSalidadeMercanciaPorCliente.Show()
            Case 9 'Ventas por Vendedor 
                Me.IsMdiContainer = True
                frmVtasRPTVentasSalidadeMercanciaPorVendedor.MdiParent = Me
                frmVtasRPTVentasSalidadeMercanciaPorVendedor.Show()
            Case 10 'Comisiones por Vendedor 
                Me.IsMdiContainer = True
                frmVtasRPTVentasSalidadeMercanciaComisionVendedor.MdiParent = Me
                frmVtasRPTVentasSalidadeMercanciaComisionVendedor.Show()
            Case 11 'Listado de Ventas x Cliente  
                Me.IsMdiContainer = True
                frmVtasRPTVtasSalMciaListadoVtasxCte.MdiParent = Me
                frmVtasRPTVtasSalMciaListadoVtasxCte.Show()
        End Select

    End Sub

    Public Sub mnuVentasSalMerOpcRepEjec_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVentasSalMerOpcRepEjec.Click
        Dim Index As Integer = mnuVentasSalMerOpcRepEjec.GetIndex(eventSender)
        'Menu Ventas Reportes Ejecutivos
        Select Case Index
            Case 0 'Ventas y Existencias por Proveedor 
                Me.IsMdiContainer = True
                frmVtasVentasyExistenciasporProveedor.MdiParent = Me
                frmVtasVentasyExistenciasporProveedor.Show()
            Case 1 'Ventas por Resurtir 
                Me.IsMdiContainer = True
                frmVtasVentasporResurtir.MdiParent = Me
                frmVtasVentasporResurtir.Show()
            Case 2 'Ventas y Utilidad Global por Grupo 
                Me.IsMdiContainer = True
                frmVtasVentasyUtilidad.MdiParent = Me
                frmVtasVentasyUtilidad.Show()
            Case 3 'Ventas por Grupo 
                Me.IsMdiContainer = True
                frmVentasPorGrupo.MdiParent = Me
                frmVentasPorGrupo.Show()
            Case 4 'Utilidad por Grupo 
                Me.IsMdiContainer = True
                frmUtilidadporGrupo.MdiParent = Me
                frmUtilidadporGrupo.Show()
            Case 5 'Ventas y Existencia por Familia 
                Me.IsMdiContainer = True
                frmVtasVentasyExistxFam.MdiParent = Me
                frmVtasVentasyExistxFam.Show()
        End Select
    End Sub

    Public Sub mnuVentasVendExtOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVentasVendExtOpc.Click
        Dim Index As Integer = mnuVentasVendExtOpc.GetIndex(eventSender)
        'Menu: Ventas - Vendedor Externo
        Select Case Index
            Case 0 'Entrada de Mercancía 
                Me.IsMdiContainer = True
                frmVtasVEEntradadeMercancia.MdiParent = Me
                frmVtasVEEntradadeMercancia.Show()
            Case 1 'salida de Mercancía 
                Me.IsMdiContainer = True
                frmVtasVESalidadeMercancia.MdiParent = Me
                frmVtasVESalidadeMercancia.Show()
            Case 2 'Reporte de Existencias con Precio Público o al Costo 
                Me.IsMdiContainer = True
                frmVtasVEExistencias.MdiParent = Me
                frmVtasVEExistencias.Show()
            Case 3 'Liquidación 
                Me.IsMdiContainer = True
                frmVtasVELiquidacionVendedorExterno.MdiParent = Me
                frmVtasVELiquidacionVendedorExterno.Show()
            Case 4 'Ingresos por Salida de Mercancía 
                Me.IsMdiContainer = True
                frmVtasVEIngresosSalidadeMercanciaaVendExt.MdiParent = Me
                frmVtasVEIngresosSalidadeMercanciaaVendExt.Show()
            Case 5 'Reporte Detallado de Entradas/Salidas 
                Me.IsMdiContainer = True
                frmVtasVEDetalladodeEntradasSalidas.MdiParent = Me
                frmVtasVEDetalladodeEntradasSalidas.Show()
        End Select
    End Sub

    Public Sub mnuVentasVtasIngrOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVentasVtasIngrOpc.Click
        Dim Index As Integer = mnuVentasVtasIngrOpc.GetIndex(eventSender)
        'Menu: Ventas - Ventas ingresos
        Select Case Index
            Case 0 'Ingresos Generales 
                Me.IsMdiContainer = True
                frmVtasRPTIngresosGenerales.MdiParent = Me
                frmVtasRPTIngresosGenerales.Show()
            Case 1 'Por Periodo y Tienda 
                Me.IsMdiContainer = True
                frmVtasRPTIngresosPorPeriodoySucursal.MdiParent = Me
                frmVtasRPTIngresosPorPeriodoySucursal.Show()
            Case 2 'Abonos 
                Me.IsMdiContainer = True
                frmVtasRPTIngresosPorAbonos.MdiParent = Me
                frmVtasRPTIngresosPorAbonos.Show()
            Case 3 'Por Reparación 
                Me.IsMdiContainer = True
                frmVtasRPTIngresosPorReparaciones.MdiParent = Me
                frmVtasRPTIngresosPorReparaciones.Show()
            Case 4 'Ingresos por Concepto de Pago 
                Me.IsMdiContainer = True
                frmVtasRPTIngresosPorConceptoDePago.MdiParent = Me
                frmVtasRPTIngresosPorConceptoDePago.Show()
        End Select
    End Sub

    '    Public Sub mnuVerToolBarOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVerToolBarOpc.Click
    '        Dim Index As Integer = mnuVerToolBarOpc.GetIndex(eventSender)
    '        mnuVerToolBarOpc(Index).Checked = Not (mnuVerToolBarOpc(Index).Checked)
    '        Select Case Index
    '            Case 0 'Ver Barra de Herramientas Estándar
    '                If mnuVerToolBarOpc(0).Checked Then
    '                    ToolbarStandar.Visible = True
    '                Else
    '                    ToolbarStandar.Visible = False
    '                End If
    '            Case 1 'Ver Barra de Estado
    '                If mnuVerToolBarOpc(1).Checked Then
    '                    Status.Visible = True
    '                Else
    '                    Status.Visible = False
    '                End If
    '        End Select
    '    End Sub

    '    Public Sub mnuVerVentanaOpc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVerVentanaOpc.Click
    '        Dim Index As Integer = mnuVerVentanaOpc.GetIndex(eventSender)
    '        Select Case Index
    '            Case 0 'Horizontal
    '                Me.LayoutMdi(System.Windows.Forms.MdiLayout.TileHorizontal)
    '            Case 1 'Vertical
    '                Me.LayoutMdi(System.Windows.Forms.MdiLayout.TileVertical)
    '            Case 2 'Cascada
    '                Me.LayoutMdi(System.Windows.Forms.MdiLayout.Cascade)
    '            Case 3 'Organizar iconos
    '                Me.LayoutMdi(System.Windows.Forms.MdiLayout.ArrangeIcons)
    '        End Select
    '    End Sub

    '    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
    '        AcumuladoTimer = AcumuladoTimer + 1

    '        If AcumuladoTimer = gintLapsoDifStock Then
    '            ValidarStockSucursales()
    '            Call GenerarPagosVirtuales()
    '            AcumuladoTimer = 0
    '        End If
    '    End Sub

    '    Private Sub GenerarPagosVirtuales()
    '        On Error GoTo Merr
    '        Dim blnTransaction As Boolean
    '        Dim I As Integer
    '        Dim J As Integer
    '        Dim rsLocal As ADODB.Recordset
    '        Dim dFechaPago As Date
    '        Dim dFechaCorte As Date

    '        dFechaCorte = Today '+ 5

    '        gStrSql = " SELECT a.FolioProgramacionP, a.CodProvAcreed, a.TipoFacturaCxP, a.TipoGasto, " & " a.FolioFactura, a.FechaFactura, a.FechaPago, a.TotalPago, a.Moneda, a.TipoCambio, a.TipoCambioE, " & " a.DescuentoFinanciero, a.SubtotalDF, a.IvaDF, a.Estatus, a.FechaCancel, a.TipoPagoProg, " & " a.Efectivo, a.Frecuencia, a.TipoIntervalo, a.Repeticiones, a.FechaInicio, a.FechaFin, a.Periodo, " & " a.DiaSemana, a.DiaMes, a.Mes, a.Opcion, a.Cual, a.Cuando " & " FROM PPIntervalos a " & " WHERE a.TipoIntervalo = 0 and a.Estatus = 'V' and a.FechaFin < '" & VB6.Format(dFechaCorte, C_FORMATFECHAGUARDAR) & "'"
    '        ModEstandar.BorraCmd()
    '        Cmd.CommandText = "dbo.UP_Select_Datos"
    '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
    '        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
    '        rsLocal = Cmd.Execute

    '        If rsLocal.RecordCount > 0 Then
    '            rsLocal.MoveFirst()
    '            Cnn.BeginTrans()
    '            blnTransaction = True
    '            For I = 1 To rsLocal.RecordCount
    '                Call ModFrecuencia.GenerarFrecuencia(rsLocal.Fields("Frecuencia").Value, rsLocal.Fields("TipoIntervalo").Value, rsLocal.Fields("Repeticiones").Value, (rsLocal.Fields("FechaFin").Value + 1), dFechaCorte, rsLocal.Fields("Periodo").Value, rsLocal.Fields("DiaSemana").Value, rsLocal.Fields("DiaMes").Value, rsLocal.Fields("Mes").Value, rsLocal.Fields("Opcion").Value, rsLocal.Fields("Cual").Value, rsLocal.Fields("Cuando").Value)
    '                'Tomar el vector del módulo ModFrecuencia, e insertar un renglón en la tabla ProgramacionPagos
    '                'por cada elemento
    '                If ModFrecuencia.nRepeticiones > 0 Then
    '                    For J = 1 To ModFrecuencia.nRepeticiones
    '                        dFechaPago = ModFrecuencia.aFechasFrecuencia(J)
    '                        'Añadir el pago a la tabla ProgramacionPagos
    '                        ModStoredProcedures.PR_IMEProgramacionPagos(Trim(rsLocal.Fields("FolioProgramacionP").Value), "0", CStr(rsLocal.Fields("CodProvAcreed").Value), Trim(rsLocal.Fields("TipoFacturaCxP").Value), Trim(rsLocal.Fields("TipoGasto").Value), Trim(rsLocal.Fields("FolioFactura").Value), VB6.Format(rsLocal.Fields("FechaFactura").Value, C_FORMATFECHAGUARDAR), VB6.Format(dFechaPago, C_FORMATFECHAGUARDAR), CStr(rsLocal.Fields("TotalPago").Value), Trim(rsLocal.Fields("Moneda").Value), CStr(rsLocal.Fields("TipoCambio").Value), CStr(rsLocal.Fields("TipoCambioE").Value), CStr(rsLocal.Fields("DescuentoFinanciero").Value), CStr(rsLocal.Fields("SubTotalDF").Value), CStr(rsLocal.Fields("IvaDF").Value), Trim(rsLocal.Fields("Estatus").Value), VB6.Format(rsLocal.Fields("FechaCancel").Value, C_FORMATFECHAGUARDAR), CStr(rsLocal.Fields("TipoPagoProg").Value), CStr(rsLocal.Fields("Efectivo").Value), "0", "1/1/1900", C_INSERCION, CStr(1))
    '                        Cmd.Execute()
    '                    Next J
    '                End If
    '                'Actualizar la fecha final y el número de repeticiones en la tabla PPIntervalos
    '                ModStoredProcedures.PR_IMEPPIntervalos(Trim(rsLocal.Fields("FolioProgramacionP").Value), "0", "", "", "", "1/1/1900", "1/1/1900", "0", "", "0", "0", "0", "0", "0", "", "1/1/1900", "", "False", "", "0", rsLocal.Fields("Repeticiones").Value + ModFrecuencia.nRepeticiones, "1/1/1900", VB6.Format(dFechaPago, C_FORMATFECHAGUARDAR), "0", "", "0", "0", "0", "0", "0", C_MODIFICACION, CStr(3))
    '                Cmd.Execute()
    '                rsLocal.MoveNext()
    '            Next I
    '            Cnn.CommitTrans()
    '            blnTransaction = False
    '        End If
    'Merr:
    '        If Err.Number <> 0 Then
    '            If blnTransaction Then Cnn.RollbackTrans()
    '            ModEstandar.MostrarError()
    '        End If
    '    End Sub

    'Private Sub ToolbarStandar_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _ToolbarStandar_Button1.Click, _ToolbarStandar_Button2.Click, _ToolbarStandar_Button3.Click, _ToolbarStandar_Button4.Click, _ToolbarStandar_Button5.Click, _ToolbarStandar_Button6.Click, _ToolbarStandar_Button7.Click
    '    Dim Button As System.Windows.Forms.ToolStripItem = CType(eventSender, System.Windows.Forms.ToolStripItem)
    '    On Error Resume Next
    '    Select Case Button.Owner.Items.IndexOf(Button)
    '        Case 1 'Guardar
    '            'ActiveMdiChild.Guardar()
    '        Case 2 'Imprimir
    '            'ActiveMdiChild.Imprime()
    '        Case 4 'Nuevo
    '            'ActiveMdiChild.Limpiar()
    '        Case 5 'Cancelar
    '            'ActiveMdiChild.Cancelar()
    '        Case 6 'Eliminar
    '            'ActiveMdiChild.Eliminar()
    '        Case 7 'Buscar
    '            'ActiveMdiChild.Buscar()
    '    End Select
    'End Sub

    Public Sub Salir()
        Me.Close()
    End Sub

    Private Function ConfiguracionCorpo() As Boolean
        On Error GoTo Merr
        GETCODGRUPOS()
        If Not ModCorporativo.CargarDatosConfiguracionCorpo Then Exit Function
        If Not ModCorporativo.CargarDatosConfiguracionPV Then Exit Function
        If Not ModCorporativo.CreaCarpetaInformes Then Exit Function
        ConfiguracionCorpo = True
        Exit Function

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Private Sub MDIMenuPrincipalCorpo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeComponent()
        lbluser.Text = gStrNomUsuario
        'btnmin.Visible = False
        'btnmin.Enabled = False
        btnmaxi.Visible = False
        btnmaxi.Enabled = False
        CargarDatosConfiguracionCorpo()
        Me.WindowState = FormWindowState.Maximized

        'ModEstandar.CentrarForma(Me)

        'If Not ConfiguracionCorpo() Then
        '    Me.Close()
        '    End
        'End If

        'gbytCantidadDecimales = 2
        'gstrFormatoCantidad = "##,##0." & New String("0", gbytCantidadDecimales)



        On Error GoTo Errores

        'Status.Items.Item(6).Text = VB6.Format(Now, "dd/MMM/yyyy")
        'DesActivar la Barra de Menu
        ModCorporativo.DesHabilitaMenuPrincipal()
        'If Not ConfiguracionCorpo() Then
        '    Me.Close()
        '    End
        'End If

        'Activar todas las opciones del Menu principal
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModCorporativo.CargarRutaImpresoras()
        gbytCantidadDecimales = 2
        gstrFormatoCantidad = "##,##0." & New String("0", gbytCantidadDecimales)
        AcumuladoTimer = 0

        If ModCorporativo.FE Then
            Me.Close()
            Exit Sub
        End If
        'FrmAcceso.Show()

Errores:
        If Err.Number <> 0 Then
            ModErrores.Errores()
            Err.Clear()
            End
        End If

        'ToolbarStandar.Visible = False
        'ToolbarStandar.BackColor = System.Windows.Forms.MenuStrip.DefaultBackColor  
        'Me.Size() = New System.Drawing.Size(1200, 600)

        'For Each file As FileInfo In dirPath.GetFiles("*.png", SearchOption.AllDirectories)
        '    ListImage.Add(Image.FromFile(file.FullName))
        'Next
        'For Each file As FileInfo In dirPath.GetFiles("*.jpg", SearchOption.AllDirectories)
        '    ListImage.Add(Image.FromFile(file.FullName))
        'Next
        'For Each file As FileInfo In dirPath.GetFiles("*.bmp", SearchOption.AllDirectories)
        '    ListImage.Add(Image.FromFile(file.FullName))
        'Next

        'MessageBox.Show("Image count " & ListImage.Count)
        'ScaleImage(PictureBox1, bit)
        'PictureBox1.Size = New System.Drawing.Size(980, 720)
        'PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage


        'TIEMPO DE CAMBIO DE LAS IMAGENES DE FONDO DEL MENU 1 MINUTO = 100000
        'Timer1.Interval = 1000
        'Timer1.Start()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'PictureBox1.Image = ListImage(r.Next(0, ListImage.Count))
        'Me.BackgroundImage = ListImage(r.Next(0, ListImage.Count))
        lblhora.Text = TimeOfDay
        lblfecha.Text = Today
    End Sub

    Private Sub btncerrar_Click(sender As Object, e As EventArgs) Handles btncerrar.Click
        Me.Close()
        System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub btnmin_Click(sender As Object, e As EventArgs) Handles btnmin.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub btnmaxi_Click(sender As Object, e As EventArgs) Handles btnmaxi.Click
        Me.WindowState = FormWindowState.Maximized
    End Sub

    'Private Sub ScaleImage(ByVal p As PictureBox, ByRef i As Bitmap)
    '    If i.Height > p.Height Then
    '        Dim diff As Integer = i.Height - p.Height
    '        Dim Resized As Bitmap = New Bitmap(i, New Size(i.Width - diff, i.Height - diff))
    '        i = Resized
    '    End If
    '    If i.Width > p.Width Then
    '        Dim diff As Integer = i.Width - p.Width
    '        Dim Resized As Bitmap = New Bitmap(i, New Size(i.Width - diff, i.Height - diff))
    '        i = Resized
    '    End If
    'End Sub
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnSoporte = New System.Windows.Forms.Button()
        Me.ButtonTeleMarketing = New System.Windows.Forms.Button()
        Me.ButtonCorteDiario = New System.Windows.Forms.Button()
        Me.ButtonRegistroCobranza = New System.Windows.Forms.Button()
        Me.ButtonRegistroGastos = New System.Windows.Forms.Button()
        Me.ButtonConsultaInventario = New System.Windows.Forms.Button()
        Me.ButtonCompraEmergencia = New System.Windows.Forms.Button()
        Me.ButtonSalidasAOrden = New System.Windows.Forms.Button()
        Me.ButtonRecepcionProducto = New System.Windows.Forms.Button()
        Me.ButtonOrdenCompra = New System.Windows.Forms.Button()
        Me.ButtonHistorial = New System.Windows.Forms.Button()
        Me.ButtonCalendario = New System.Windows.Forms.Button()
        Me.ButtonCalculadora = New System.Windows.Forms.Button()
        Me.ButtonCotizacion = New System.Windows.Forms.Button()
        Me.ButtonEmpresas = New System.Windows.Forms.Button()
        Me.ButtonClientes = New System.Windows.Forms.Button()
        Me.ButtonRecepcion = New System.Windows.Forms.Button()
        Me.ButtonPanelControl = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.imgEstandar = New System.Windows.Forms.ImageList(Me.components)
        Me.MenuAcercaDe = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._MenuAcercaDe_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuContextualGenOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._menuContextualGenOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._menuContextualGenOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._menuContextualGenOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._menuContextualGenOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._menuContextualGenOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._menuContextualGenOpc_6 = New System.Windows.Forms.ToolStripMenuItem()
        Me._menuContextualGenOpc_7 = New System.Windows.Forms.ToolStripMenuItem()
        Me._menuContextualGenOpc_8 = New System.Windows.Forms.ToolStripMenuItem()
        Me._menuContextualGenOpc_9 = New System.Windows.Forms.ToolStripMenuItem()
        Me._menuContextualGenOpc_10 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuArchivoOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuBancosOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuBancosOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiario_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiario_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiario_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiario_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiario_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiario_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiario_6 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiario_7 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiario_8 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiario_9 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiarioRptOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiarioRptOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoDiarioRptOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoMensual_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoMensual_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoMensual_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoMensual_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoMensual_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcProcesoMensual_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcCatalogos_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcCatalogos_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcCatalogos_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuBancosOpcCatalogos_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuBancosOpcCatalogos = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuBancosOpcProcesoDiario = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuBancosOpcProcesoDiarioRptOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuBancosOpcProcesoMensual = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuCatalogosOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuCatalogosOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_6 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_7 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_8 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_9 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_10 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_11 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_12 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_13 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_14 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_16 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_17 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_19 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_20 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_22 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_23 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_24 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCompyCxPOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuCompyCxPOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPOpc_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPOpc_6 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPOpc_7 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPRptOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPRptOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPRptOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPRptOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPRptOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPRptOpc_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPRptOpc_6 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuCompyCxPRptOpc_7 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPRptOpc_8 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPRptOpc_9 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPRptOpc_10 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPOpc_9 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCompyCxPRptOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuConfiguracionOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuConfiguracionOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuConfiguracionOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuConfiguracionOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuConfiguracionOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuConfiguracionOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuConfiguracionOpc_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuConfiguracionOpc_7 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuConfiguracionOpc_9 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuConfiguracionOpc_10 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuConfiguracionOpc_12 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuConfiguracionOpcUtil_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuConfiguracionOpcUtil = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuContextualOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuEdicionOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuFacturacionOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuFacturacionOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFacturacionOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFacturacionOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFacturacionRptFactOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFacturacionRptFactOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFacturacionRptFactOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFacturacionRptFactOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFacturacionRptFactOpc_4 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuFacturacionRptFactOpc_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFacturacionOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFacturacionOpc_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFacturacionRptFactOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuInvEntradasOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuInvEntradasOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvEntradasOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvEntradasOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvEntradasOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvEntradasOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvEntradasOpc_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvEntradasOpc_6 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvEntradasOpc_7 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInvHojaOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuInvHojaOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvHojaOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInvOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuInvOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvSalidasOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvSalidasOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvSalidasOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvSalidasOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvSalidasOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvSalidasOpc_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvSalidasOpc_6 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvSalidasOpc_7 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvOpc_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvRptOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvRptOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvRptOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvRptOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuInvRptOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInvRptOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuInvSalidasOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuSegOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuSegOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuSegOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuSegOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuVentasOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me._mnuVentasOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_6 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_7 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_8 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_9 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_10 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_11 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpc_12 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuVentasSalMerOpc_13 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpcRepEjec_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpcRepEjec_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpcRepEjec_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpcRepEjec_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpcRepEjec_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasSalMerOpcRepEjec_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasVtasIngrOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasVtasIngrOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasVtasIngrOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasVtasIngrOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasVtasIngrOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasVendExtOpc_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasVendExtOpc_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasVendExtOpc_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasVendExtOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasVendExtOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasVendExtOpc_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasOpc_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasOpc_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasOpc_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasOpc_7 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasOpc_8 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasOpc_10 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasOpc_11 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuVentasSalMerOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuVentasSalMerOpcRepEjec = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuVentasVendExtOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuVentasVtasIngrOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuVerOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuVerToolBarOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.mnuVerVentanaOpc = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuCatalogos = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCatalogosOpc_15 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuCatalogosOpc_18 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuCatalogosOpc_21 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuVentas = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuVentasOpc_6 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuVentasOpc_9 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuComprasyCxP = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuCompyCxPOpc_8 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuFacturacion = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuFacturacionOpc_3 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuBancos = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInventarios = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuConfiguracion = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuConfiguracionOpc_6 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuConfiguracionOpc_8 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuConfiguracionOpc_11 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuSeg = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuContextualGen = New System.Windows.Forms.ToolStripMenuItem()
        Me._menuContextualGenOpc_5 = New System.Windows.Forms.ToolStripSeparator()
        Me.ButtonContainer = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblActualizacion = New System.Windows.Forms.Label()
        Me.panel2 = New System.Windows.Forms.Panel()
        Me.btnmaxi = New System.Windows.Forms.Button()
        Me.lblhora = New System.Windows.Forms.Label()
        Me.lblfecha = New System.Windows.Forms.Label()
        Me.lbluser = New System.Windows.Forms.Label()
        Me.btnmin = New System.Windows.Forms.Button()
        Me.btncerrar = New System.Windows.Forms.Button()
        CType(Me.MenuAcercaDe, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.menuContextualGenOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuArchivoOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuBancosOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuBancosOpcCatalogos, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuBancosOpcProcesoDiario, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuBancosOpcProcesoDiarioRptOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuBancosOpcProcesoMensual, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuCatalogosOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuCompyCxPOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuCompyCxPRptOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuConfiguracionOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuConfiguracionOpcUtil, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuContextualOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuEdicionOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuFacturacionOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuFacturacionRptFactOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuInvEntradasOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuInvHojaOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuInvOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuInvRptOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuInvSalidasOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuSegOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuVentasOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuVentasSalMerOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuVentasSalMerOpcRepEjec, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuVentasVendExtOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuVentasVtasIngrOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuVerOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuVerToolBarOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuVerVentanaOpc, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainMenu1.SuspendLayout()
        Me.ButtonContainer.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnSoporte
        '
        Me.btnSoporte.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.btnSoporte.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSoporte.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnSoporte.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSoporte.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.0!)
        Me.btnSoporte.ForeColor = System.Drawing.Color.Black
        Me.btnSoporte.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnSoporte.Location = New System.Drawing.Point(1114, 0)
        Me.btnSoporte.Name = "btnSoporte"
        Me.btnSoporte.Size = New System.Drawing.Size(66, 71)
        Me.btnSoporte.TabIndex = 217
        Me.btnSoporte.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.btnSoporte, "Generacion de Ticket para solicitud de Soporte")
        Me.btnSoporte.UseVisualStyleBackColor = False
        '
        'ButtonTeleMarketing
        '
        Me.ButtonTeleMarketing.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonTeleMarketing.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonTeleMarketing.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonTeleMarketing.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonTeleMarketing.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.0!)
        Me.ButtonTeleMarketing.ForeColor = System.Drawing.Color.Black
        Me.ButtonTeleMarketing.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonTeleMarketing.Location = New System.Drawing.Point(1048, 0)
        Me.ButtonTeleMarketing.Name = "ButtonTeleMarketing"
        Me.ButtonTeleMarketing.Size = New System.Drawing.Size(66, 71)
        Me.ButtonTeleMarketing.TabIndex = 214
        Me.ButtonTeleMarketing.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonTeleMarketing, "CRM")
        Me.ButtonTeleMarketing.UseVisualStyleBackColor = False
        '
        'ButtonCorteDiario
        '
        Me.ButtonCorteDiario.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonCorteDiario.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonCorteDiario.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonCorteDiario.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonCorteDiario.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.0!)
        Me.ButtonCorteDiario.ForeColor = System.Drawing.Color.Black
        Me.ButtonCorteDiario.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonCorteDiario.Location = New System.Drawing.Point(982, 0)
        Me.ButtonCorteDiario.Name = "ButtonCorteDiario"
        Me.ButtonCorteDiario.Size = New System.Drawing.Size(66, 71)
        Me.ButtonCorteDiario.TabIndex = 212
        Me.ButtonCorteDiario.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonCorteDiario, "Corte Diario")
        Me.ButtonCorteDiario.UseVisualStyleBackColor = False
        '
        'ButtonRegistroCobranza
        '
        Me.ButtonRegistroCobranza.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonRegistroCobranza.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonRegistroCobranza.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonRegistroCobranza.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonRegistroCobranza.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.0!)
        Me.ButtonRegistroCobranza.ForeColor = System.Drawing.Color.Black
        Me.ButtonRegistroCobranza.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonRegistroCobranza.Location = New System.Drawing.Point(916, 0)
        Me.ButtonRegistroCobranza.Name = "ButtonRegistroCobranza"
        Me.ButtonRegistroCobranza.Size = New System.Drawing.Size(66, 71)
        Me.ButtonRegistroCobranza.TabIndex = 211
        Me.ButtonRegistroCobranza.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonRegistroCobranza, "Registro de Cobranza")
        Me.ButtonRegistroCobranza.UseVisualStyleBackColor = False
        '
        'ButtonRegistroGastos
        '
        Me.ButtonRegistroGastos.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonRegistroGastos.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonRegistroGastos.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonRegistroGastos.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonRegistroGastos.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.0!)
        Me.ButtonRegistroGastos.ForeColor = System.Drawing.Color.Black
        Me.ButtonRegistroGastos.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonRegistroGastos.Location = New System.Drawing.Point(850, 0)
        Me.ButtonRegistroGastos.Name = "ButtonRegistroGastos"
        Me.ButtonRegistroGastos.Size = New System.Drawing.Size(66, 71)
        Me.ButtonRegistroGastos.TabIndex = 210
        Me.ButtonRegistroGastos.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonRegistroGastos, "Registro de Gastos")
        Me.ButtonRegistroGastos.UseVisualStyleBackColor = False
        '
        'ButtonConsultaInventario
        '
        Me.ButtonConsultaInventario.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonConsultaInventario.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonConsultaInventario.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonConsultaInventario.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonConsultaInventario.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.0!)
        Me.ButtonConsultaInventario.ForeColor = System.Drawing.Color.Black
        Me.ButtonConsultaInventario.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonConsultaInventario.Location = New System.Drawing.Point(792, 0)
        Me.ButtonConsultaInventario.Name = "ButtonConsultaInventario"
        Me.ButtonConsultaInventario.Size = New System.Drawing.Size(58, 71)
        Me.ButtonConsultaInventario.TabIndex = 209
        Me.ButtonConsultaInventario.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonConsultaInventario, "Consulta de Inventario")
        Me.ButtonConsultaInventario.UseVisualStyleBackColor = False
        '
        'ButtonCompraEmergencia
        '
        Me.ButtonCompraEmergencia.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonCompraEmergencia.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonCompraEmergencia.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonCompraEmergencia.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonCompraEmergencia.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.0!)
        Me.ButtonCompraEmergencia.ForeColor = System.Drawing.Color.Black
        Me.ButtonCompraEmergencia.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonCompraEmergencia.Location = New System.Drawing.Point(726, 0)
        Me.ButtonCompraEmergencia.Name = "ButtonCompraEmergencia"
        Me.ButtonCompraEmergencia.Size = New System.Drawing.Size(66, 71)
        Me.ButtonCompraEmergencia.TabIndex = 208
        Me.ButtonCompraEmergencia.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonCompraEmergencia, "Consulta de Puntos Midas")
        Me.ButtonCompraEmergencia.UseVisualStyleBackColor = False
        '
        'ButtonSalidasAOrden
        '
        Me.ButtonSalidasAOrden.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonSalidasAOrden.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonSalidasAOrden.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonSalidasAOrden.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonSalidasAOrden.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.0!)
        Me.ButtonSalidasAOrden.ForeColor = System.Drawing.Color.Black
        Me.ButtonSalidasAOrden.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonSalidasAOrden.Location = New System.Drawing.Point(660, 0)
        Me.ButtonSalidasAOrden.Name = "ButtonSalidasAOrden"
        Me.ButtonSalidasAOrden.Size = New System.Drawing.Size(66, 71)
        Me.ButtonSalidasAOrden.TabIndex = 207
        Me.ButtonSalidasAOrden.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonSalidasAOrden, "Salidas a Ordenes de Servicio")
        Me.ButtonSalidasAOrden.UseVisualStyleBackColor = False
        '
        'ButtonRecepcionProducto
        '
        Me.ButtonRecepcionProducto.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonRecepcionProducto.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonRecepcionProducto.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonRecepcionProducto.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonRecepcionProducto.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.0!)
        Me.ButtonRecepcionProducto.ForeColor = System.Drawing.Color.Black
        Me.ButtonRecepcionProducto.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonRecepcionProducto.Location = New System.Drawing.Point(594, 0)
        Me.ButtonRecepcionProducto.Name = "ButtonRecepcionProducto"
        Me.ButtonRecepcionProducto.Size = New System.Drawing.Size(66, 71)
        Me.ButtonRecepcionProducto.TabIndex = 206
        Me.ButtonRecepcionProducto.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonRecepcionProducto, "Recepción de Producto")
        Me.ButtonRecepcionProducto.UseVisualStyleBackColor = False
        '
        'ButtonOrdenCompra
        '
        Me.ButtonOrdenCompra.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonOrdenCompra.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonOrdenCompra.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonOrdenCompra.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonOrdenCompra.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.0!)
        Me.ButtonOrdenCompra.ForeColor = System.Drawing.Color.Black
        Me.ButtonOrdenCompra.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonOrdenCompra.Location = New System.Drawing.Point(528, 0)
        Me.ButtonOrdenCompra.Name = "ButtonOrdenCompra"
        Me.ButtonOrdenCompra.Size = New System.Drawing.Size(66, 71)
        Me.ButtonOrdenCompra.TabIndex = 205
        Me.ButtonOrdenCompra.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonOrdenCompra, "Orden de Compra")
        Me.ButtonOrdenCompra.UseVisualStyleBackColor = False
        '
        'ButtonHistorial
        '
        Me.ButtonHistorial.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonHistorial.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonHistorial.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonHistorial.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonHistorial.Font = New System.Drawing.Font("Microsoft Sans Serif", 5.0!)
        Me.ButtonHistorial.ForeColor = System.Drawing.Color.Black
        Me.ButtonHistorial.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonHistorial.Location = New System.Drawing.Point(462, 0)
        Me.ButtonHistorial.Name = "ButtonHistorial"
        Me.ButtonHistorial.Size = New System.Drawing.Size(66, 71)
        Me.ButtonHistorial.TabIndex = 204
        Me.ButtonHistorial.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonHistorial, "Historial del Cliente")
        Me.ButtonHistorial.UseVisualStyleBackColor = False
        '
        'ButtonCalendario
        '
        Me.ButtonCalendario.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonCalendario.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonCalendario.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonCalendario.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonCalendario.Font = New System.Drawing.Font("Microsoft Sans Serif", 5.0!)
        Me.ButtonCalendario.ForeColor = System.Drawing.Color.Black
        Me.ButtonCalendario.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonCalendario.Location = New System.Drawing.Point(396, 0)
        Me.ButtonCalendario.Name = "ButtonCalendario"
        Me.ButtonCalendario.Size = New System.Drawing.Size(66, 71)
        Me.ButtonCalendario.TabIndex = 203
        Me.ButtonCalendario.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonCalendario, "Calendario/ Citas")
        Me.ButtonCalendario.UseVisualStyleBackColor = False
        '
        'ButtonCalculadora
        '
        Me.ButtonCalculadora.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonCalculadora.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonCalculadora.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonCalculadora.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonCalculadora.Font = New System.Drawing.Font("Microsoft Sans Serif", 5.0!)
        Me.ButtonCalculadora.ForeColor = System.Drawing.Color.Black
        Me.ButtonCalculadora.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonCalculadora.Location = New System.Drawing.Point(330, 0)
        Me.ButtonCalculadora.Name = "ButtonCalculadora"
        Me.ButtonCalculadora.Size = New System.Drawing.Size(66, 71)
        Me.ButtonCalculadora.TabIndex = 23
        Me.ButtonCalculadora.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonCalculadora, "Calculadora de Precios")
        Me.ButtonCalculadora.UseVisualStyleBackColor = False
        '
        'ButtonCotizacion
        '
        Me.ButtonCotizacion.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonCotizacion.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonCotizacion.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonCotizacion.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonCotizacion.Font = New System.Drawing.Font("Microsoft Sans Serif", 5.0!)
        Me.ButtonCotizacion.ForeColor = System.Drawing.Color.Black
        Me.ButtonCotizacion.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonCotizacion.Location = New System.Drawing.Point(264, 0)
        Me.ButtonCotizacion.Name = "ButtonCotizacion"
        Me.ButtonCotizacion.Size = New System.Drawing.Size(66, 71)
        Me.ButtonCotizacion.TabIndex = 22
        Me.ButtonCotizacion.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonCotizacion, "Cotización Rápida")
        Me.ButtonCotizacion.UseVisualStyleBackColor = False
        '
        'ButtonEmpresas
        '
        Me.ButtonEmpresas.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonEmpresas.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonEmpresas.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonEmpresas.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonEmpresas.Font = New System.Drawing.Font("Microsoft Sans Serif", 5.0!)
        Me.ButtonEmpresas.ForeColor = System.Drawing.Color.Black
        Me.ButtonEmpresas.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonEmpresas.Location = New System.Drawing.Point(198, 0)
        Me.ButtonEmpresas.Name = "ButtonEmpresas"
        Me.ButtonEmpresas.Size = New System.Drawing.Size(66, 71)
        Me.ButtonEmpresas.TabIndex = 21
        Me.ButtonEmpresas.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonEmpresas, "Empresas")
        Me.ButtonEmpresas.UseVisualStyleBackColor = False
        '
        'ButtonClientes
        '
        Me.ButtonClientes.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonClientes.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonClientes.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonClientes.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonClientes.Font = New System.Drawing.Font("Microsoft Sans Serif", 5.0!)
        Me.ButtonClientes.ForeColor = System.Drawing.Color.Black
        Me.ButtonClientes.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonClientes.Location = New System.Drawing.Point(132, 0)
        Me.ButtonClientes.Name = "ButtonClientes"
        Me.ButtonClientes.Size = New System.Drawing.Size(66, 71)
        Me.ButtonClientes.TabIndex = 17
        Me.ButtonClientes.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonClientes, "Clientes")
        Me.ButtonClientes.UseVisualStyleBackColor = False
        '
        'ButtonRecepcion
        '
        Me.ButtonRecepcion.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonRecepcion.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonRecepcion.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonRecepcion.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonRecepcion.Font = New System.Drawing.Font("Microsoft Sans Serif", 5.0!)
        Me.ButtonRecepcion.ForeColor = System.Drawing.Color.Black
        Me.ButtonRecepcion.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonRecepcion.Location = New System.Drawing.Point(66, 0)
        Me.ButtonRecepcion.Name = "ButtonRecepcion"
        Me.ButtonRecepcion.Size = New System.Drawing.Size(66, 71)
        Me.ButtonRecepcion.TabIndex = 15
        Me.ButtonRecepcion.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonRecepcion, "Recepción de Clientes")
        Me.ButtonRecepcion.UseVisualStyleBackColor = False
        '
        'ButtonPanelControl
        '
        Me.ButtonPanelControl.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonPanelControl.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ButtonPanelControl.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonPanelControl.Dock = System.Windows.Forms.DockStyle.Left
        Me.ButtonPanelControl.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonPanelControl.Font = New System.Drawing.Font("Microsoft Sans Serif", 5.0!)
        Me.ButtonPanelControl.ForeColor = System.Drawing.Color.Black
        Me.ButtonPanelControl.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ButtonPanelControl.Location = New System.Drawing.Point(0, 0)
        Me.ButtonPanelControl.Name = "ButtonPanelControl"
        Me.ButtonPanelControl.Size = New System.Drawing.Size(66, 71)
        Me.ButtonPanelControl.TabIndex = 14
        Me.ButtonPanelControl.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.ButtonPanelControl, "Panel de Control")
        Me.ButtonPanelControl.UseVisualStyleBackColor = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 900
        '
        'imgEstandar
        '
        Me.imgEstandar.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.imgEstandar.ImageSize = New System.Drawing.Size(16, 16)
        Me.imgEstandar.TransparentColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        '
        'MenuAcercaDe
        '
        '
        '_MenuAcercaDe_0
        '
        Me._MenuAcercaDe_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.MenuAcercaDe.SetIndex(Me._MenuAcercaDe_0, CType(0, Short))
        Me._MenuAcercaDe_0.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me._MenuAcercaDe_0.Name = "_MenuAcercaDe_0"
        Me._MenuAcercaDe_0.Size = New System.Drawing.Size(80, 20)
        Me._MenuAcercaDe_0.Text = "Acerca &de..."
        '
        '_menuContextualGenOpc_0
        '
        Me._menuContextualGenOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.menuContextualGenOpc.SetIndex(Me._menuContextualGenOpc_0, CType(0, Short))
        Me._menuContextualGenOpc_0.Name = "_menuContextualGenOpc_0"
        Me._menuContextualGenOpc_0.Size = New System.Drawing.Size(125, 22)
        Me._menuContextualGenOpc_0.Text = "ARCHIVO"
        '
        '_menuContextualGenOpc_1
        '
        Me._menuContextualGenOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.menuContextualGenOpc.SetIndex(Me._menuContextualGenOpc_1, CType(1, Short))
        Me._menuContextualGenOpc_1.Name = "_menuContextualGenOpc_1"
        Me._menuContextualGenOpc_1.Size = New System.Drawing.Size(125, 22)
        Me._menuContextualGenOpc_1.Text = "Guardar"
        '
        '_menuContextualGenOpc_2
        '
        Me._menuContextualGenOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.menuContextualGenOpc.SetIndex(Me._menuContextualGenOpc_2, CType(2, Short))
        Me._menuContextualGenOpc_2.Name = "_menuContextualGenOpc_2"
        Me._menuContextualGenOpc_2.Size = New System.Drawing.Size(125, 22)
        Me._menuContextualGenOpc_2.Text = "Imprimir"
        '
        '_menuContextualGenOpc_3
        '
        Me._menuContextualGenOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.menuContextualGenOpc.SetIndex(Me._menuContextualGenOpc_3, CType(3, Short))
        Me._menuContextualGenOpc_3.Name = "_menuContextualGenOpc_3"
        Me._menuContextualGenOpc_3.Size = New System.Drawing.Size(125, 22)
        Me._menuContextualGenOpc_3.Text = "Cerrar"
        '
        '_menuContextualGenOpc_4
        '
        Me._menuContextualGenOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.menuContextualGenOpc.SetIndex(Me._menuContextualGenOpc_4, CType(4, Short))
        Me._menuContextualGenOpc_4.Name = "_menuContextualGenOpc_4"
        Me._menuContextualGenOpc_4.Size = New System.Drawing.Size(125, 22)
        Me._menuContextualGenOpc_4.Text = "Salir"
        '
        '_menuContextualGenOpc_6
        '
        Me._menuContextualGenOpc_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.menuContextualGenOpc.SetIndex(Me._menuContextualGenOpc_6, CType(6, Short))
        Me._menuContextualGenOpc_6.Name = "_menuContextualGenOpc_6"
        Me._menuContextualGenOpc_6.Size = New System.Drawing.Size(125, 22)
        Me._menuContextualGenOpc_6.Text = "EDICIÓN"
        '
        '_menuContextualGenOpc_7
        '
        Me._menuContextualGenOpc_7.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.menuContextualGenOpc.SetIndex(Me._menuContextualGenOpc_7, CType(7, Short))
        Me._menuContextualGenOpc_7.Name = "_menuContextualGenOpc_7"
        Me._menuContextualGenOpc_7.Size = New System.Drawing.Size(125, 22)
        Me._menuContextualGenOpc_7.Text = "Nuevo"
        '
        '_menuContextualGenOpc_8
        '
        Me._menuContextualGenOpc_8.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.menuContextualGenOpc.SetIndex(Me._menuContextualGenOpc_8, CType(8, Short))
        Me._menuContextualGenOpc_8.Name = "_menuContextualGenOpc_8"
        Me._menuContextualGenOpc_8.Size = New System.Drawing.Size(125, 22)
        Me._menuContextualGenOpc_8.Text = "Cancelar"
        '
        '_menuContextualGenOpc_9
        '
        Me._menuContextualGenOpc_9.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.menuContextualGenOpc.SetIndex(Me._menuContextualGenOpc_9, CType(9, Short))
        Me._menuContextualGenOpc_9.Name = "_menuContextualGenOpc_9"
        Me._menuContextualGenOpc_9.Size = New System.Drawing.Size(125, 22)
        Me._menuContextualGenOpc_9.Text = "Eliminar"
        '
        '_menuContextualGenOpc_10
        '
        Me._menuContextualGenOpc_10.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.menuContextualGenOpc.SetIndex(Me._menuContextualGenOpc_10, CType(10, Short))
        Me._menuContextualGenOpc_10.Name = "_menuContextualGenOpc_10"
        Me._menuContextualGenOpc_10.Size = New System.Drawing.Size(125, 22)
        Me._menuContextualGenOpc_10.Text = "Buscar"
        '
        'mnuArchivoOpc
        '
        '
        'mnuBancosOpc
        '
        '
        '_mnuBancosOpc_0
        '
        Me._mnuBancosOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuBancosOpc_0.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuBancosOpcProcesoDiario_0, Me._mnuBancosOpcProcesoDiario_1, Me._mnuBancosOpcProcesoDiario_2, Me._mnuBancosOpcProcesoDiario_3, Me._mnuBancosOpcProcesoDiario_4, Me._mnuBancosOpcProcesoDiario_5, Me._mnuBancosOpcProcesoDiario_6, Me._mnuBancosOpcProcesoDiario_7, Me._mnuBancosOpcProcesoDiario_8, Me._mnuBancosOpcProcesoDiario_9})
        Me.mnuBancosOpc.SetIndex(Me._mnuBancosOpc_0, CType(0, Short))
        Me._mnuBancosOpc_0.Name = "_mnuBancosOpc_0"
        Me._mnuBancosOpc_0.Size = New System.Drawing.Size(280, 22)
        Me._mnuBancosOpc_0.Text = "Proceso Diario"
        '
        '_mnuBancosOpcProcesoDiario_0
        '
        Me._mnuBancosOpcProcesoDiario_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoDiario.SetIndex(Me._mnuBancosOpcProcesoDiario_0, CType(0, Short))
        Me._mnuBancosOpcProcesoDiario_0.Name = "_mnuBancosOpcProcesoDiario_0"
        Me._mnuBancosOpcProcesoDiario_0.Size = New System.Drawing.Size(196, 22)
        Me._mnuBancosOpcProcesoDiario_0.Text = "Registro de Pagos"
        '
        '_mnuBancosOpcProcesoDiario_1
        '
        Me._mnuBancosOpcProcesoDiario_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoDiario.SetIndex(Me._mnuBancosOpcProcesoDiario_1, CType(1, Short))
        Me._mnuBancosOpcProcesoDiario_1.Name = "_mnuBancosOpcProcesoDiario_1"
        Me._mnuBancosOpcProcesoDiario_1.Size = New System.Drawing.Size(196, 22)
        Me._mnuBancosOpcProcesoDiario_1.Text = "Depósitos"
        '
        '_mnuBancosOpcProcesoDiario_2
        '
        Me._mnuBancosOpcProcesoDiario_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoDiario.SetIndex(Me._mnuBancosOpcProcesoDiario_2, CType(2, Short))
        Me._mnuBancosOpcProcesoDiario_2.Name = "_mnuBancosOpcProcesoDiario_2"
        Me._mnuBancosOpcProcesoDiario_2.Size = New System.Drawing.Size(196, 22)
        Me._mnuBancosOpcProcesoDiario_2.Text = "Cargos Diversos"
        '
        '_mnuBancosOpcProcesoDiario_3
        '
        Me._mnuBancosOpcProcesoDiario_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoDiario.SetIndex(Me._mnuBancosOpcProcesoDiario_3, CType(3, Short))
        Me._mnuBancosOpcProcesoDiario_3.Name = "_mnuBancosOpcProcesoDiario_3"
        Me._mnuBancosOpcProcesoDiario_3.Size = New System.Drawing.Size(196, 22)
        Me._mnuBancosOpcProcesoDiario_3.Text = "Traspasos Bancarios"
        '
        '_mnuBancosOpcProcesoDiario_4
        '
        Me._mnuBancosOpcProcesoDiario_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoDiario.SetIndex(Me._mnuBancosOpcProcesoDiario_4, CType(4, Short))
        Me._mnuBancosOpcProcesoDiario_4.Name = "_mnuBancosOpcProcesoDiario_4"
        Me._mnuBancosOpcProcesoDiario_4.Size = New System.Drawing.Size(196, 22)
        Me._mnuBancosOpcProcesoDiario_4.Text = "Anticipo a Proveedores"
        '
        '_mnuBancosOpcProcesoDiario_5
        '
        Me._mnuBancosOpcProcesoDiario_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoDiario.SetIndex(Me._mnuBancosOpcProcesoDiario_5, CType(5, Short))
        Me._mnuBancosOpcProcesoDiario_5.Name = "_mnuBancosOpcProcesoDiario_5"
        Me._mnuBancosOpcProcesoDiario_5.Size = New System.Drawing.Size(196, 22)
        Me._mnuBancosOpcProcesoDiario_5.Text = "Otros Ingresos"
        '
        '_mnuBancosOpcProcesoDiario_6
        '
        Me._mnuBancosOpcProcesoDiario_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoDiario.SetIndex(Me._mnuBancosOpcProcesoDiario_6, CType(6, Short))
        Me._mnuBancosOpcProcesoDiario_6.Name = "_mnuBancosOpcProcesoDiario_6"
        Me._mnuBancosOpcProcesoDiario_6.Size = New System.Drawing.Size(196, 22)
        Me._mnuBancosOpcProcesoDiario_6.Text = "Cancelar Movimientos"
        '
        '_mnuBancosOpcProcesoDiario_7
        '
        Me._mnuBancosOpcProcesoDiario_7.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoDiario.SetIndex(Me._mnuBancosOpcProcesoDiario_7, CType(7, Short))
        Me._mnuBancosOpcProcesoDiario_7.Name = "_mnuBancosOpcProcesoDiario_7"
        Me._mnuBancosOpcProcesoDiario_7.Size = New System.Drawing.Size(196, 22)
        Me._mnuBancosOpcProcesoDiario_7.Text = "Consulta de Saldos"
        '
        '_mnuBancosOpcProcesoDiario_8
        '
        Me._mnuBancosOpcProcesoDiario_8.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoDiario.SetIndex(Me._mnuBancosOpcProcesoDiario_8, CType(8, Short))
        Me._mnuBancosOpcProcesoDiario_8.Name = "_mnuBancosOpcProcesoDiario_8"
        Me._mnuBancosOpcProcesoDiario_8.Size = New System.Drawing.Size(196, 22)
        Me._mnuBancosOpcProcesoDiario_8.Text = "Cierre Diario"
        '
        '_mnuBancosOpcProcesoDiario_9
        '
        Me._mnuBancosOpcProcesoDiario_9.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuBancosOpcProcesoDiario_9.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuBancosOpcProcesoDiarioRptOpc_0, Me._mnuBancosOpcProcesoDiarioRptOpc_1, Me._mnuBancosOpcProcesoDiarioRptOpc_2})
        Me.mnuBancosOpcProcesoDiario.SetIndex(Me._mnuBancosOpcProcesoDiario_9, CType(9, Short))
        Me._mnuBancosOpcProcesoDiario_9.Name = "_mnuBancosOpcProcesoDiario_9"
        Me._mnuBancosOpcProcesoDiario_9.Size = New System.Drawing.Size(196, 22)
        Me._mnuBancosOpcProcesoDiario_9.Text = "Reportes"
        '
        '_mnuBancosOpcProcesoDiarioRptOpc_0
        '
        Me.mnuBancosOpcProcesoDiarioRptOpc.SetIndex(Me._mnuBancosOpcProcesoDiarioRptOpc_0, CType(0, Short))
        Me._mnuBancosOpcProcesoDiarioRptOpc_0.Name = "_mnuBancosOpcProcesoDiarioRptOpc_0"
        Me._mnuBancosOpcProcesoDiarioRptOpc_0.Size = New System.Drawing.Size(246, 22)
        Me._mnuBancosOpcProcesoDiarioRptOpc_0.Text = "Movimientos Bancarios"
        '
        '_mnuBancosOpcProcesoDiarioRptOpc_1
        '
        Me.mnuBancosOpcProcesoDiarioRptOpc.SetIndex(Me._mnuBancosOpcProcesoDiarioRptOpc_1, CType(1, Short))
        Me._mnuBancosOpcProcesoDiarioRptOpc_1.Name = "_mnuBancosOpcProcesoDiarioRptOpc_1"
        Me._mnuBancosOpcProcesoDiarioRptOpc_1.Size = New System.Drawing.Size(246, 22)
        Me._mnuBancosOpcProcesoDiarioRptOpc_1.Text = "Movimientos Bancarios por Tipo"
        '
        '_mnuBancosOpcProcesoDiarioRptOpc_2
        '
        Me.mnuBancosOpcProcesoDiarioRptOpc.SetIndex(Me._mnuBancosOpcProcesoDiarioRptOpc_2, CType(2, Short))
        Me._mnuBancosOpcProcesoDiarioRptOpc_2.Name = "_mnuBancosOpcProcesoDiarioRptOpc_2"
        Me._mnuBancosOpcProcesoDiarioRptOpc_2.Size = New System.Drawing.Size(246, 22)
        Me._mnuBancosOpcProcesoDiarioRptOpc_2.Text = "Analisis Diario de Bancos"
        '
        '_mnuBancosOpc_1
        '
        Me._mnuBancosOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuBancosOpc_1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuBancosOpcProcesoMensual_0, Me._mnuBancosOpcProcesoMensual_1, Me._mnuBancosOpcProcesoMensual_2, Me._mnuBancosOpcProcesoMensual_3, Me._mnuBancosOpcProcesoMensual_4, Me._mnuBancosOpcProcesoMensual_5})
        Me.mnuBancosOpc.SetIndex(Me._mnuBancosOpc_1, CType(1, Short))
        Me._mnuBancosOpc_1.Name = "_mnuBancosOpc_1"
        Me._mnuBancosOpc_1.Size = New System.Drawing.Size(280, 22)
        Me._mnuBancosOpc_1.Text = "Proceso Mensual"
        '
        '_mnuBancosOpcProcesoMensual_0
        '
        Me._mnuBancosOpcProcesoMensual_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoMensual.SetIndex(Me._mnuBancosOpcProcesoMensual_0, CType(0, Short))
        Me._mnuBancosOpcProcesoMensual_0.Name = "_mnuBancosOpcProcesoMensual_0"
        Me._mnuBancosOpcProcesoMensual_0.Size = New System.Drawing.Size(310, 22)
        Me._mnuBancosOpcProcesoMensual_0.Text = "Conciliación Mensual"
        '
        '_mnuBancosOpcProcesoMensual_1
        '
        Me._mnuBancosOpcProcesoMensual_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoMensual.SetIndex(Me._mnuBancosOpcProcesoMensual_1, CType(1, Short))
        Me._mnuBancosOpcProcesoMensual_1.Name = "_mnuBancosOpcProcesoMensual_1"
        Me._mnuBancosOpcProcesoMensual_1.Size = New System.Drawing.Size(310, 22)
        Me._mnuBancosOpcProcesoMensual_1.Text = "Reportes de Movimientos en Conciliación"
        '
        '_mnuBancosOpcProcesoMensual_2
        '
        Me._mnuBancosOpcProcesoMensual_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoMensual.SetIndex(Me._mnuBancosOpcProcesoMensual_2, CType(2, Short))
        Me._mnuBancosOpcProcesoMensual_2.Name = "_mnuBancosOpcProcesoMensual_2"
        Me._mnuBancosOpcProcesoMensual_2.Size = New System.Drawing.Size(310, 22)
        Me._mnuBancosOpcProcesoMensual_2.Text = "Flujo de la Caja General"
        '
        '_mnuBancosOpcProcesoMensual_3
        '
        Me._mnuBancosOpcProcesoMensual_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoMensual.SetIndex(Me._mnuBancosOpcProcesoMensual_3, CType(3, Short))
        Me._mnuBancosOpcProcesoMensual_3.Name = "_mnuBancosOpcProcesoMensual_3"
        Me._mnuBancosOpcProcesoMensual_3.Size = New System.Drawing.Size(310, 22)
        Me._mnuBancosOpcProcesoMensual_3.Text = "Consulta de Origen y Aplicación de Recursos"
        '
        '_mnuBancosOpcProcesoMensual_4
        '
        Me._mnuBancosOpcProcesoMensual_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoMensual.SetIndex(Me._mnuBancosOpcProcesoMensual_4, CType(4, Short))
        Me._mnuBancosOpcProcesoMensual_4.Name = "_mnuBancosOpcProcesoMensual_4"
        Me._mnuBancosOpcProcesoMensual_4.Size = New System.Drawing.Size(310, 22)
        Me._mnuBancosOpcProcesoMensual_4.Text = "Estado de Origen y Aplicación de Recursos"
        '
        '_mnuBancosOpcProcesoMensual_5
        '
        Me._mnuBancosOpcProcesoMensual_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcProcesoMensual.SetIndex(Me._mnuBancosOpcProcesoMensual_5, CType(5, Short))
        Me._mnuBancosOpcProcesoMensual_5.Name = "_mnuBancosOpcProcesoMensual_5"
        Me._mnuBancosOpcProcesoMensual_5.Size = New System.Drawing.Size(310, 22)
        Me._mnuBancosOpcProcesoMensual_5.Text = "Cierre de Conciliación Mensual"
        '
        '_mnuBancosOpc_2
        '
        Me._mnuBancosOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpc.SetIndex(Me._mnuBancosOpc_2, CType(2, Short))
        Me._mnuBancosOpc_2.Name = "_mnuBancosOpc_2"
        Me._mnuBancosOpc_2.Size = New System.Drawing.Size(280, 22)
        Me._mnuBancosOpc_2.Text = "Depuración de Movimientos Historicos"
        '
        '_mnuBancosOpc_3
        '
        Me._mnuBancosOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuBancosOpc_3.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuBancosOpcCatalogos_0, Me._mnuBancosOpcCatalogos_1, Me._mnuBancosOpcCatalogos_2, Me._mnuBancosOpcCatalogos_3})
        Me.mnuBancosOpc.SetIndex(Me._mnuBancosOpc_3, CType(3, Short))
        Me._mnuBancosOpc_3.Name = "_mnuBancosOpc_3"
        Me._mnuBancosOpc_3.Size = New System.Drawing.Size(280, 22)
        Me._mnuBancosOpc_3.Text = "Catálogos"
        '
        '_mnuBancosOpcCatalogos_0
        '
        Me._mnuBancosOpcCatalogos_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcCatalogos.SetIndex(Me._mnuBancosOpcCatalogos_0, CType(0, Short))
        Me._mnuBancosOpcCatalogos_0.Name = "_mnuBancosOpcCatalogos_0"
        Me._mnuBancosOpcCatalogos_0.Size = New System.Drawing.Size(270, 22)
        Me._mnuBancosOpcCatalogos_0.Text = "ABC Bancos"
        '
        '_mnuBancosOpcCatalogos_1
        '
        Me._mnuBancosOpcCatalogos_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcCatalogos.SetIndex(Me._mnuBancosOpcCatalogos_1, CType(1, Short))
        Me._mnuBancosOpcCatalogos_1.Name = "_mnuBancosOpcCatalogos_1"
        Me._mnuBancosOpcCatalogos_1.Size = New System.Drawing.Size(270, 22)
        Me._mnuBancosOpcCatalogos_1.Text = "ABC Cuentas Bancarias"
        '
        '_mnuBancosOpcCatalogos_2
        '
        Me._mnuBancosOpcCatalogos_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcCatalogos.SetIndex(Me._mnuBancosOpcCatalogos_2, CType(2, Short))
        Me._mnuBancosOpcCatalogos_2.Name = "_mnuBancosOpcCatalogos_2"
        Me._mnuBancosOpcCatalogos_2.Size = New System.Drawing.Size(270, 22)
        Me._mnuBancosOpcCatalogos_2.Text = "ABC Origen y Aplicación de Recursos"
        '
        '_mnuBancosOpcCatalogos_3
        '
        Me._mnuBancosOpcCatalogos_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuBancosOpcCatalogos.SetIndex(Me._mnuBancosOpcCatalogos_3, CType(3, Short))
        Me._mnuBancosOpcCatalogos_3.Name = "_mnuBancosOpcCatalogos_3"
        Me._mnuBancosOpcCatalogos_3.Size = New System.Drawing.Size(270, 22)
        Me._mnuBancosOpcCatalogos_3.Text = "ABC Rubros de Aplicación y Origen"
        '
        'mnuBancosOpcCatalogos
        '
        '
        'mnuBancosOpcProcesoDiario
        '
        '
        'mnuBancosOpcProcesoDiarioRptOpc
        '
        '
        'mnuBancosOpcProcesoMensual
        '
        '
        'mnuCatalogosOpc
        '
        '
        '_mnuCatalogosOpc_0
        '
        Me._mnuCatalogosOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_0, CType(0, Short))
        Me._mnuCatalogosOpc_0.Name = "_mnuCatalogosOpc_0"
        Me._mnuCatalogosOpc_0.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_0.Text = "Clientes"
        '
        '_mnuCatalogosOpc_1
        '
        Me._mnuCatalogosOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_1, CType(1, Short))
        Me._mnuCatalogosOpc_1.Name = "_mnuCatalogosOpc_1"
        Me._mnuCatalogosOpc_1.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_1.Text = "Vendedores"
        '
        '_mnuCatalogosOpc_2
        '
        Me._mnuCatalogosOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_2, CType(2, Short))
        Me._mnuCatalogosOpc_2.Name = "_mnuCatalogosOpc_2"
        Me._mnuCatalogosOpc_2.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_2.Text = "Sucursales"
        '
        '_mnuCatalogosOpc_3
        '
        Me._mnuCatalogosOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_3, CType(3, Short))
        Me._mnuCatalogosOpc_3.Name = "_mnuCatalogosOpc_3"
        Me._mnuCatalogosOpc_3.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_3.Text = "Talleres"
        '
        '_mnuCatalogosOpc_4
        '
        Me._mnuCatalogosOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_4, CType(4, Short))
        Me._mnuCatalogosOpc_4.Name = "_mnuCatalogosOpc_4"
        Me._mnuCatalogosOpc_4.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_4.Text = "Tipos de Material"
        '
        '_mnuCatalogosOpc_5
        '
        Me._mnuCatalogosOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_5, CType(5, Short))
        Me._mnuCatalogosOpc_5.Name = "_mnuCatalogosOpc_5"
        Me._mnuCatalogosOpc_5.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_5.Text = "Proveedores y Acreedores"
        '
        '_mnuCatalogosOpc_6
        '
        Me._mnuCatalogosOpc_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_6, CType(6, Short))
        Me._mnuCatalogosOpc_6.Name = "_mnuCatalogosOpc_6"
        Me._mnuCatalogosOpc_6.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_6.Text = "Formas de Pago"
        '
        '_mnuCatalogosOpc_7
        '
        Me._mnuCatalogosOpc_7.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_7, CType(7, Short))
        Me._mnuCatalogosOpc_7.Name = "_mnuCatalogosOpc_7"
        Me._mnuCatalogosOpc_7.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_7.Text = "Bancos"
        '
        '_mnuCatalogosOpc_8
        '
        Me._mnuCatalogosOpc_8.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_8, CType(8, Short))
        Me._mnuCatalogosOpc_8.Name = "_mnuCatalogosOpc_8"
        Me._mnuCatalogosOpc_8.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_8.Text = "Cuentas Bancarias"
        '
        '_mnuCatalogosOpc_9
        '
        Me._mnuCatalogosOpc_9.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_9, CType(9, Short))
        Me._mnuCatalogosOpc_9.Name = "_mnuCatalogosOpc_9"
        Me._mnuCatalogosOpc_9.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_9.Text = "Origen y Aplicación de Recursos"
        '
        '_mnuCatalogosOpc_10
        '
        Me._mnuCatalogosOpc_10.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_10, CType(10, Short))
        Me._mnuCatalogosOpc_10.Name = "_mnuCatalogosOpc_10"
        Me._mnuCatalogosOpc_10.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_10.Text = "Rubros de Origen y Aplicación"
        '
        '_mnuCatalogosOpc_11
        '
        Me._mnuCatalogosOpc_11.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_11, CType(11, Short))
        Me._mnuCatalogosOpc_11.Name = "_mnuCatalogosOpc_11"
        Me._mnuCatalogosOpc_11.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_11.Text = "Descuentos a Vendedores Externos"
        '
        '_mnuCatalogosOpc_12
        '
        Me._mnuCatalogosOpc_12.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_12, CType(12, Short))
        Me._mnuCatalogosOpc_12.Name = "_mnuCatalogosOpc_12"
        Me._mnuCatalogosOpc_12.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_12.Text = "Promociones de Tarjetas Bancarias"
        '
        '_mnuCatalogosOpc_13
        '
        Me._mnuCatalogosOpc_13.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_13, CType(13, Short))
        Me._mnuCatalogosOpc_13.Name = "_mnuCatalogosOpc_13"
        Me._mnuCatalogosOpc_13.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_13.Text = "Programación de Promociones"
        '
        '_mnuCatalogosOpc_14
        '
        Me._mnuCatalogosOpc_14.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_14, CType(14, Short))
        Me._mnuCatalogosOpc_14.Name = "_mnuCatalogosOpc_14"
        Me._mnuCatalogosOpc_14.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_14.Text = "Comisiones por  ventas de vendedores"
        '
        '_mnuCatalogosOpc_16
        '
        Me._mnuCatalogosOpc_16.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_16, CType(16, Short))
        Me._mnuCatalogosOpc_16.Name = "_mnuCatalogosOpc_16"
        Me._mnuCatalogosOpc_16.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_16.Text = "Artículos"
        '
        '_mnuCatalogosOpc_17
        '
        Me._mnuCatalogosOpc_17.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_17, CType(17, Short))
        Me._mnuCatalogosOpc_17.Name = "_mnuCatalogosOpc_17"
        Me._mnuCatalogosOpc_17.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_17.Text = "Grupos de Artículos"
        Me._mnuCatalogosOpc_17.Visible = False
        '
        '_mnuCatalogosOpc_19
        '
        Me._mnuCatalogosOpc_19.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_19, CType(19, Short))
        Me._mnuCatalogosOpc_19.Name = "_mnuCatalogosOpc_19"
        Me._mnuCatalogosOpc_19.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_19.Text = "Marcas de Relojería"
        '
        '_mnuCatalogosOpc_20
        '
        Me._mnuCatalogosOpc_20.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_20, CType(20, Short))
        Me._mnuCatalogosOpc_20.Name = "_mnuCatalogosOpc_20"
        Me._mnuCatalogosOpc_20.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_20.Text = "Modelos de Relojería"
        '
        '_mnuCatalogosOpc_22
        '
        Me._mnuCatalogosOpc_22.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_22, CType(22, Short))
        Me._mnuCatalogosOpc_22.Name = "_mnuCatalogosOpc_22"
        Me._mnuCatalogosOpc_22.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_22.Text = "Familias de Joyería y Varios"
        '
        '_mnuCatalogosOpc_23
        '
        Me._mnuCatalogosOpc_23.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_23, CType(23, Short))
        Me._mnuCatalogosOpc_23.Name = "_mnuCatalogosOpc_23"
        Me._mnuCatalogosOpc_23.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_23.Text = "Líneas de Joyería y Varios"
        '
        '_mnuCatalogosOpc_24
        '
        Me._mnuCatalogosOpc_24.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCatalogosOpc.SetIndex(Me._mnuCatalogosOpc_24, CType(24, Short))
        Me._mnuCatalogosOpc_24.Name = "_mnuCatalogosOpc_24"
        Me._mnuCatalogosOpc_24.Size = New System.Drawing.Size(277, 22)
        Me._mnuCatalogosOpc_24.Text = "Sub Líneas de Joyería"
        '
        'mnuCompyCxPOpc
        '
        '
        '_mnuCompyCxPOpc_0
        '
        Me._mnuCompyCxPOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPOpc.SetIndex(Me._mnuCompyCxPOpc_0, CType(0, Short))
        Me._mnuCompyCxPOpc_0.Name = "_mnuCompyCxPOpc_0"
        Me._mnuCompyCxPOpc_0.Size = New System.Drawing.Size(247, 22)
        Me._mnuCompyCxPOpc_0.Text = "Orden de Compra"
        '
        '_mnuCompyCxPOpc_1
        '
        Me._mnuCompyCxPOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPOpc.SetIndex(Me._mnuCompyCxPOpc_1, CType(1, Short))
        Me._mnuCompyCxPOpc_1.Name = "_mnuCompyCxPOpc_1"
        Me._mnuCompyCxPOpc_1.Size = New System.Drawing.Size(247, 22)
        Me._mnuCompyCxPOpc_1.Text = "Registro de Facturas de Compras"
        '
        '_mnuCompyCxPOpc_2
        '
        Me._mnuCompyCxPOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPOpc.SetIndex(Me._mnuCompyCxPOpc_2, CType(2, Short))
        Me._mnuCompyCxPOpc_2.Name = "_mnuCompyCxPOpc_2"
        Me._mnuCompyCxPOpc_2.Size = New System.Drawing.Size(247, 22)
        Me._mnuCompyCxPOpc_2.Text = "Registro de Facturas de Gastos"
        '
        '_mnuCompyCxPOpc_3
        '
        Me._mnuCompyCxPOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPOpc.SetIndex(Me._mnuCompyCxPOpc_3, CType(3, Short))
        Me._mnuCompyCxPOpc_3.Name = "_mnuCompyCxPOpc_3"
        Me._mnuCompyCxPOpc_3.Size = New System.Drawing.Size(247, 22)
        Me._mnuCompyCxPOpc_3.Text = "Programación Especial de Pagos"
        '
        '_mnuCompyCxPOpc_4
        '
        Me._mnuCompyCxPOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPOpc.SetIndex(Me._mnuCompyCxPOpc_4, CType(4, Short))
        Me._mnuCompyCxPOpc_4.Name = "_mnuCompyCxPOpc_4"
        Me._mnuCompyCxPOpc_4.Size = New System.Drawing.Size(247, 22)
        Me._mnuCompyCxPOpc_4.Text = "Notas de Crédito"
        '
        '_mnuCompyCxPOpc_5
        '
        Me._mnuCompyCxPOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPOpc.SetIndex(Me._mnuCompyCxPOpc_5, CType(5, Short))
        Me._mnuCompyCxPOpc_5.Name = "_mnuCompyCxPOpc_5"
        Me._mnuCompyCxPOpc_5.Size = New System.Drawing.Size(247, 22)
        Me._mnuCompyCxPOpc_5.Text = "Cuentas por Pagar"
        Me._mnuCompyCxPOpc_5.Visible = False
        '
        '_mnuCompyCxPOpc_6
        '
        Me._mnuCompyCxPOpc_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPOpc.SetIndex(Me._mnuCompyCxPOpc_6, CType(6, Short))
        Me._mnuCompyCxPOpc_6.Name = "_mnuCompyCxPOpc_6"
        Me._mnuCompyCxPOpc_6.Size = New System.Drawing.Size(247, 22)
        Me._mnuCompyCxPOpc_6.Text = "Emisión de pagos"
        '
        '_mnuCompyCxPOpc_7
        '
        Me._mnuCompyCxPOpc_7.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuCompyCxPOpc_7.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuCompyCxPRptOpc_0, Me._mnuCompyCxPRptOpc_1, Me._mnuCompyCxPRptOpc_2, Me._mnuCompyCxPRptOpc_3, Me._mnuCompyCxPRptOpc_4, Me._mnuCompyCxPRptOpc_5, Me._mnuCompyCxPRptOpc_6, Me._mnuCompyCxPRptOpc_7, Me._mnuCompyCxPRptOpc_8, Me._mnuCompyCxPRptOpc_9, Me._mnuCompyCxPRptOpc_10})
        Me.mnuCompyCxPOpc.SetIndex(Me._mnuCompyCxPOpc_7, CType(7, Short))
        Me._mnuCompyCxPOpc_7.Name = "_mnuCompyCxPOpc_7"
        Me._mnuCompyCxPOpc_7.Size = New System.Drawing.Size(247, 22)
        Me._mnuCompyCxPOpc_7.Text = "Reportes"
        '
        '_mnuCompyCxPRptOpc_0
        '
        Me._mnuCompyCxPRptOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPRptOpc.SetIndex(Me._mnuCompyCxPRptOpc_0, CType(0, Short))
        Me._mnuCompyCxPRptOpc_0.Name = "_mnuCompyCxPRptOpc_0"
        Me._mnuCompyCxPRptOpc_0.Size = New System.Drawing.Size(242, 22)
        Me._mnuCompyCxPRptOpc_0.Text = "Reporte de Facturas"
        '
        '_mnuCompyCxPRptOpc_1
        '
        Me._mnuCompyCxPRptOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPRptOpc.SetIndex(Me._mnuCompyCxPRptOpc_1, CType(1, Short))
        Me._mnuCompyCxPRptOpc_1.Name = "_mnuCompyCxPRptOpc_1"
        Me._mnuCompyCxPRptOpc_1.Size = New System.Drawing.Size(242, 22)
        Me._mnuCompyCxPRptOpc_1.Text = "Reporte de Notas de Crédito"
        '
        '_mnuCompyCxPRptOpc_2
        '
        Me._mnuCompyCxPRptOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPRptOpc.SetIndex(Me._mnuCompyCxPRptOpc_2, CType(2, Short))
        Me._mnuCompyCxPRptOpc_2.Name = "_mnuCompyCxPRptOpc_2"
        Me._mnuCompyCxPRptOpc_2.Size = New System.Drawing.Size(242, 22)
        Me._mnuCompyCxPRptOpc_2.Text = "Reporte de los Mejores"
        '
        '_mnuCompyCxPRptOpc_3
        '
        Me._mnuCompyCxPRptOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPRptOpc.SetIndex(Me._mnuCompyCxPRptOpc_3, CType(3, Short))
        Me._mnuCompyCxPRptOpc_3.Name = "_mnuCompyCxPRptOpc_3"
        Me._mnuCompyCxPRptOpc_3.Size = New System.Drawing.Size(242, 22)
        Me._mnuCompyCxPRptOpc_3.Text = "Análisis Anual de Compras"
        '
        '_mnuCompyCxPRptOpc_4
        '
        Me._mnuCompyCxPRptOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPRptOpc.SetIndex(Me._mnuCompyCxPRptOpc_4, CType(4, Short))
        Me._mnuCompyCxPRptOpc_4.Name = "_mnuCompyCxPRptOpc_4"
        Me._mnuCompyCxPRptOpc_4.Size = New System.Drawing.Size(242, 22)
        Me._mnuCompyCxPRptOpc_4.Text = "Órdenes de Compra"
        '
        '_mnuCompyCxPRptOpc_5
        '
        Me._mnuCompyCxPRptOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPRptOpc.SetIndex(Me._mnuCompyCxPRptOpc_5, CType(5, Short))
        Me._mnuCompyCxPRptOpc_5.Name = "_mnuCompyCxPRptOpc_5"
        Me._mnuCompyCxPRptOpc_5.Size = New System.Drawing.Size(242, 22)
        Me._mnuCompyCxPRptOpc_5.Text = "Artículos Pendientes por Recibir"
        '
        '_mnuCompyCxPRptOpc_6
        '
        Me._mnuCompyCxPRptOpc_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuCompyCxPRptOpc_6.Name = "_mnuCompyCxPRptOpc_6"
        Me._mnuCompyCxPRptOpc_6.Size = New System.Drawing.Size(239, 6)
        '
        '_mnuCompyCxPRptOpc_7
        '
        Me._mnuCompyCxPRptOpc_7.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPRptOpc.SetIndex(Me._mnuCompyCxPRptOpc_7, CType(7, Short))
        Me._mnuCompyCxPRptOpc_7.Name = "_mnuCompyCxPRptOpc_7"
        Me._mnuCompyCxPRptOpc_7.Size = New System.Drawing.Size(242, 22)
        Me._mnuCompyCxPRptOpc_7.Text = "Cuentas por Pagar"
        '
        '_mnuCompyCxPRptOpc_8
        '
        Me._mnuCompyCxPRptOpc_8.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPRptOpc.SetIndex(Me._mnuCompyCxPRptOpc_8, CType(8, Short))
        Me._mnuCompyCxPRptOpc_8.Name = "_mnuCompyCxPRptOpc_8"
        Me._mnuCompyCxPRptOpc_8.Size = New System.Drawing.Size(242, 22)
        Me._mnuCompyCxPRptOpc_8.Text = "Auxiliar de Proveedores"
        '
        '_mnuCompyCxPRptOpc_9
        '
        Me._mnuCompyCxPRptOpc_9.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPRptOpc.SetIndex(Me._mnuCompyCxPRptOpc_9, CType(9, Short))
        Me._mnuCompyCxPRptOpc_9.Name = "_mnuCompyCxPRptOpc_9"
        Me._mnuCompyCxPRptOpc_9.Size = New System.Drawing.Size(242, 22)
        Me._mnuCompyCxPRptOpc_9.Text = "CXP Presupuestado"
        '
        '_mnuCompyCxPRptOpc_10
        '
        Me._mnuCompyCxPRptOpc_10.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPRptOpc.SetIndex(Me._mnuCompyCxPRptOpc_10, CType(10, Short))
        Me._mnuCompyCxPRptOpc_10.Name = "_mnuCompyCxPRptOpc_10"
        Me._mnuCompyCxPRptOpc_10.Size = New System.Drawing.Size(242, 22)
        Me._mnuCompyCxPRptOpc_10.Text = "Saldos por Proveedores"
        '
        '_mnuCompyCxPOpc_9
        '
        Me._mnuCompyCxPOpc_9.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuCompyCxPOpc.SetIndex(Me._mnuCompyCxPOpc_9, CType(9, Short))
        Me._mnuCompyCxPOpc_9.Name = "_mnuCompyCxPOpc_9"
        Me._mnuCompyCxPOpc_9.Size = New System.Drawing.Size(247, 22)
        Me._mnuCompyCxPOpc_9.Text = "Carga Inicial Facturas"
        '
        'mnuCompyCxPRptOpc
        '
        '
        'mnuConfiguracionOpc
        '
        '
        '_mnuConfiguracionOpc_0
        '
        Me._mnuConfiguracionOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuConfiguracionOpc.SetIndex(Me._mnuConfiguracionOpc_0, CType(0, Short))
        Me._mnuConfiguracionOpc_0.Name = "_mnuConfiguracionOpc_0"
        Me._mnuConfiguracionOpc_0.Size = New System.Drawing.Size(302, 22)
        Me._mnuConfiguracionOpc_0.Text = "Configuración General"
        '
        '_mnuConfiguracionOpc_1
        '
        Me._mnuConfiguracionOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuConfiguracionOpc.SetIndex(Me._mnuConfiguracionOpc_1, CType(1, Short))
        Me._mnuConfiguracionOpc_1.Name = "_mnuConfiguracionOpc_1"
        Me._mnuConfiguracionOpc_1.Size = New System.Drawing.Size(302, 22)
        Me._mnuConfiguracionOpc_1.Text = "Configuración del Punto de Venta"
        '
        '_mnuConfiguracionOpc_2
        '
        Me._mnuConfiguracionOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuConfiguracionOpc.SetIndex(Me._mnuConfiguracionOpc_2, CType(2, Short))
        Me._mnuConfiguracionOpc_2.Name = "_mnuConfiguracionOpc_2"
        Me._mnuConfiguracionOpc_2.Size = New System.Drawing.Size(302, 22)
        Me._mnuConfiguracionOpc_2.Text = "Configuración de Tickets"
        '
        '_mnuConfiguracionOpc_3
        '
        Me._mnuConfiguracionOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuConfiguracionOpc.SetIndex(Me._mnuConfiguracionOpc_3, CType(3, Short))
        Me._mnuConfiguracionOpc_3.Name = "_mnuConfiguracionOpc_3"
        Me._mnuConfiguracionOpc_3.Size = New System.Drawing.Size(302, 22)
        Me._mnuConfiguracionOpc_3.Text = "Configuración de Facturas"
        '
        '_mnuConfiguracionOpc_4
        '
        Me._mnuConfiguracionOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuConfiguracionOpc.SetIndex(Me._mnuConfiguracionOpc_4, CType(4, Short))
        Me._mnuConfiguracionOpc_4.Name = "_mnuConfiguracionOpc_4"
        Me._mnuConfiguracionOpc_4.Size = New System.Drawing.Size(302, 22)
        Me._mnuConfiguracionOpc_4.Text = "Configuracion de Folios del Punto de Venta"
        '
        '_mnuConfiguracionOpc_5
        '
        Me._mnuConfiguracionOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuConfiguracionOpc.SetIndex(Me._mnuConfiguracionOpc_5, CType(5, Short))
        Me._mnuConfiguracionOpc_5.Name = "_mnuConfiguracionOpc_5"
        Me._mnuConfiguracionOpc_5.Size = New System.Drawing.Size(302, 22)
        Me._mnuConfiguracionOpc_5.Text = "Configuracion de Cajas del Punto de Venta"
        '
        '_mnuConfiguracionOpc_7
        '
        Me._mnuConfiguracionOpc_7.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuConfiguracionOpc.SetIndex(Me._mnuConfiguracionOpc_7, CType(7, Short))
        Me._mnuConfiguracionOpc_7.Name = "_mnuConfiguracionOpc_7"
        Me._mnuConfiguracionOpc_7.Size = New System.Drawing.Size(302, 22)
        Me._mnuConfiguracionOpc_7.Text = "Impresora"
        '
        '_mnuConfiguracionOpc_9
        '
        Me._mnuConfiguracionOpc_9.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuConfiguracionOpc.SetIndex(Me._mnuConfiguracionOpc_9, CType(9, Short))
        Me._mnuConfiguracionOpc_9.Name = "_mnuConfiguracionOpc_9"
        Me._mnuConfiguracionOpc_9.Size = New System.Drawing.Size(302, 22)
        Me._mnuConfiguracionOpc_9.Text = "Cambiar &Usuario"
        '
        '_mnuConfiguracionOpc_10
        '
        Me._mnuConfiguracionOpc_10.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuConfiguracionOpc.SetIndex(Me._mnuConfiguracionOpc_10, CType(10, Short))
        Me._mnuConfiguracionOpc_10.Name = "_mnuConfiguracionOpc_10"
        Me._mnuConfiguracionOpc_10.Size = New System.Drawing.Size(302, 22)
        Me._mnuConfiguracionOpc_10.Text = "Cambiar Contraseña"
        '
        '_mnuConfiguracionOpc_12
        '
        Me._mnuConfiguracionOpc_12.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuConfiguracionOpc_12.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuConfiguracionOpcUtil_0})
        Me.mnuConfiguracionOpc.SetIndex(Me._mnuConfiguracionOpc_12, CType(12, Short))
        Me._mnuConfiguracionOpc_12.Name = "_mnuConfiguracionOpc_12"
        Me._mnuConfiguracionOpc_12.Size = New System.Drawing.Size(302, 22)
        Me._mnuConfiguracionOpc_12.Text = "Utilerias"
        '
        '_mnuConfiguracionOpcUtil_0
        '
        Me._mnuConfiguracionOpcUtil_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuConfiguracionOpcUtil.SetIndex(Me._mnuConfiguracionOpcUtil_0, CType(0, Short))
        Me._mnuConfiguracionOpcUtil_0.Name = "_mnuConfiguracionOpcUtil_0"
        Me._mnuConfiguracionOpcUtil_0.Size = New System.Drawing.Size(174, 22)
        Me._mnuConfiguracionOpcUtil_0.Text = "Importar Imagenes"
        '
        'mnuConfiguracionOpcUtil
        '
        '
        'mnuFacturacionOpc
        '
        '
        '_mnuFacturacionOpc_0
        '
        Me._mnuFacturacionOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuFacturacionOpc.SetIndex(Me._mnuFacturacionOpc_0, CType(0, Short))
        Me._mnuFacturacionOpc_0.Name = "_mnuFacturacionOpc_0"
        Me._mnuFacturacionOpc_0.Size = New System.Drawing.Size(235, 22)
        Me._mnuFacturacionOpc_0.Text = "Análisis de las Ventas"
        '
        '_mnuFacturacionOpc_1
        '
        Me._mnuFacturacionOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuFacturacionOpc.SetIndex(Me._mnuFacturacionOpc_1, CType(1, Short))
        Me._mnuFacturacionOpc_1.Name = "_mnuFacturacionOpc_1"
        Me._mnuFacturacionOpc_1.Size = New System.Drawing.Size(235, 22)
        Me._mnuFacturacionOpc_1.Text = "Facturación Especial"
        '
        '_mnuFacturacionOpc_2
        '
        Me._mnuFacturacionOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuFacturacionOpc_2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuFacturacionRptFactOpc_0, Me._mnuFacturacionRptFactOpc_1, Me._mnuFacturacionRptFactOpc_2, Me._mnuFacturacionRptFactOpc_3, Me._mnuFacturacionRptFactOpc_4, Me._mnuFacturacionRptFactOpc_5})
        Me.mnuFacturacionOpc.SetIndex(Me._mnuFacturacionOpc_2, CType(2, Short))
        Me._mnuFacturacionOpc_2.Name = "_mnuFacturacionOpc_2"
        Me._mnuFacturacionOpc_2.Size = New System.Drawing.Size(235, 22)
        Me._mnuFacturacionOpc_2.Text = "Reportes de Facturación"
        '
        '_mnuFacturacionRptFactOpc_0
        '
        Me._mnuFacturacionRptFactOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuFacturacionRptFactOpc.SetIndex(Me._mnuFacturacionRptFactOpc_0, CType(0, Short))
        Me._mnuFacturacionRptFactOpc_0.Name = "_mnuFacturacionRptFactOpc_0"
        Me._mnuFacturacionRptFactOpc_0.Size = New System.Drawing.Size(287, 22)
        Me._mnuFacturacionRptFactOpc_0.Text = "Global por Tienda"
        '
        '_mnuFacturacionRptFactOpc_1
        '
        Me._mnuFacturacionRptFactOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuFacturacionRptFactOpc.SetIndex(Me._mnuFacturacionRptFactOpc_1, CType(1, Short))
        Me._mnuFacturacionRptFactOpc_1.Name = "_mnuFacturacionRptFactOpc_1"
        Me._mnuFacturacionRptFactOpc_1.Size = New System.Drawing.Size(287, 22)
        Me._mnuFacturacionRptFactOpc_1.Text = "Detallada por Tienda"
        '
        '_mnuFacturacionRptFactOpc_2
        '
        Me._mnuFacturacionRptFactOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuFacturacionRptFactOpc.SetIndex(Me._mnuFacturacionRptFactOpc_2, CType(2, Short))
        Me._mnuFacturacionRptFactOpc_2.Name = "_mnuFacturacionRptFactOpc_2"
        Me._mnuFacturacionRptFactOpc_2.Size = New System.Drawing.Size(287, 22)
        Me._mnuFacturacionRptFactOpc_2.Text = "Reimpresión de Tickets"
        '
        '_mnuFacturacionRptFactOpc_3
        '
        Me._mnuFacturacionRptFactOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuFacturacionRptFactOpc.SetIndex(Me._mnuFacturacionRptFactOpc_3, CType(3, Short))
        Me._mnuFacturacionRptFactOpc_3.Name = "_mnuFacturacionRptFactOpc_3"
        Me._mnuFacturacionRptFactOpc_3.Size = New System.Drawing.Size(287, 22)
        Me._mnuFacturacionRptFactOpc_3.Text = "Los Mejores Clientes"
        '
        '_mnuFacturacionRptFactOpc_4
        '
        Me._mnuFacturacionRptFactOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuFacturacionRptFactOpc_4.Name = "_mnuFacturacionRptFactOpc_4"
        Me._mnuFacturacionRptFactOpc_4.Size = New System.Drawing.Size(284, 6)
        '
        '_mnuFacturacionRptFactOpc_5
        '
        Me._mnuFacturacionRptFactOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuFacturacionRptFactOpc.SetIndex(Me._mnuFacturacionRptFactOpc_5, CType(5, Short))
        Me._mnuFacturacionRptFactOpc_5.Name = "_mnuFacturacionRptFactOpc_5"
        Me._mnuFacturacionRptFactOpc_5.Size = New System.Drawing.Size(287, 22)
        Me._mnuFacturacionRptFactOpc_5.Text = "Reporte de Ventas con Tarjeta de Credito"
        '
        '_mnuFacturacionOpc_4
        '
        Me._mnuFacturacionOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuFacturacionOpc.SetIndex(Me._mnuFacturacionOpc_4, CType(4, Short))
        Me._mnuFacturacionOpc_4.Name = "_mnuFacturacionOpc_4"
        Me._mnuFacturacionOpc_4.Size = New System.Drawing.Size(235, 22)
        Me._mnuFacturacionOpc_4.Text = "ReImpresión de Cortes de Caja"
        '
        '_mnuFacturacionOpc_5
        '
        Me._mnuFacturacionOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuFacturacionOpc.SetIndex(Me._mnuFacturacionOpc_5, CType(5, Short))
        Me._mnuFacturacionOpc_5.Name = "_mnuFacturacionOpc_5"
        Me._mnuFacturacionOpc_5.Size = New System.Drawing.Size(235, 22)
        Me._mnuFacturacionOpc_5.Text = "Diario de Movimientos"
        '
        'mnuFacturacionRptFactOpc
        '
        '
        'mnuInvEntradasOpc
        '
        '
        '_mnuInvEntradasOpc_0
        '
        Me._mnuInvEntradasOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvEntradasOpc.SetIndex(Me._mnuInvEntradasOpc_0, CType(0, Short))
        Me._mnuInvEntradasOpc_0.Name = "_mnuInvEntradasOpc_0"
        Me._mnuInvEntradasOpc_0.Size = New System.Drawing.Size(346, 22)
        Me._mnuInvEntradasOpc_0.Text = "Por Compra"
        '
        '_mnuInvEntradasOpc_1
        '
        Me._mnuInvEntradasOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvEntradasOpc.SetIndex(Me._mnuInvEntradasOpc_1, CType(1, Short))
        Me._mnuInvEntradasOpc_1.Name = "_mnuInvEntradasOpc_1"
        Me._mnuInvEntradasOpc_1.Size = New System.Drawing.Size(346, 22)
        Me._mnuInvEntradasOpc_1.Text = "Por Transferencias"
        '
        '_mnuInvEntradasOpc_2
        '
        Me._mnuInvEntradasOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvEntradasOpc.SetIndex(Me._mnuInvEntradasOpc_2, CType(2, Short))
        Me._mnuInvEntradasOpc_2.Name = "_mnuInvEntradasOpc_2"
        Me._mnuInvEntradasOpc_2.Size = New System.Drawing.Size(346, 22)
        Me._mnuInvEntradasOpc_2.Text = "Por Devolución sobre Venta"
        '
        '_mnuInvEntradasOpc_3
        '
        Me._mnuInvEntradasOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvEntradasOpc.SetIndex(Me._mnuInvEntradasOpc_3, CType(3, Short))
        Me._mnuInvEntradasOpc_3.Name = "_mnuInvEntradasOpc_3"
        Me._mnuInvEntradasOpc_3.Size = New System.Drawing.Size(346, 22)
        Me._mnuInvEntradasOpc_3.Text = "Por Devolución de Vendedores Externos"
        '
        '_mnuInvEntradasOpc_4
        '
        Me._mnuInvEntradasOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvEntradasOpc.SetIndex(Me._mnuInvEntradasOpc_4, CType(4, Short))
        Me._mnuInvEntradasOpc_4.Name = "_mnuInvEntradasOpc_4"
        Me._mnuInvEntradasOpc_4.Size = New System.Drawing.Size(346, 22)
        Me._mnuInvEntradasOpc_4.Text = "Por Devolución sobre Venta de Vendedores Externos"
        '
        '_mnuInvEntradasOpc_5
        '
        Me._mnuInvEntradasOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvEntradasOpc.SetIndex(Me._mnuInvEntradasOpc_5, CType(5, Short))
        Me._mnuInvEntradasOpc_5.Name = "_mnuInvEntradasOpc_5"
        Me._mnuInvEntradasOpc_5.Size = New System.Drawing.Size(346, 22)
        Me._mnuInvEntradasOpc_5.Text = "Devolución sobre Obsequio"
        '
        '_mnuInvEntradasOpc_6
        '
        Me._mnuInvEntradasOpc_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvEntradasOpc.SetIndex(Me._mnuInvEntradasOpc_6, CType(6, Short))
        Me._mnuInvEntradasOpc_6.Name = "_mnuInvEntradasOpc_6"
        Me._mnuInvEntradasOpc_6.Size = New System.Drawing.Size(346, 22)
        Me._mnuInvEntradasOpc_6.Text = "Devolución de Préstamo"
        '
        '_mnuInvEntradasOpc_7
        '
        Me._mnuInvEntradasOpc_7.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvEntradasOpc.SetIndex(Me._mnuInvEntradasOpc_7, CType(7, Short))
        Me._mnuInvEntradasOpc_7.Name = "_mnuInvEntradasOpc_7"
        Me._mnuInvEntradasOpc_7.Size = New System.Drawing.Size(346, 22)
        Me._mnuInvEntradasOpc_7.Text = "Por Sustitución"
        '
        'mnuInvHojaOpc
        '
        '
        '_mnuInvHojaOpc_0
        '
        Me._mnuInvHojaOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvHojaOpc.SetIndex(Me._mnuInvHojaOpc_0, CType(0, Short))
        Me._mnuInvHojaOpc_0.Name = "_mnuInvHojaOpc_0"
        Me._mnuInvHojaOpc_0.Size = New System.Drawing.Size(184, 22)
        Me._mnuInvHojaOpc_0.Text = "Hoja de control"
        '
        '_mnuInvHojaOpc_1
        '
        Me._mnuInvHojaOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvHojaOpc.SetIndex(Me._mnuInvHojaOpc_1, CType(1, Short))
        Me._mnuInvHojaOpc_1.Name = "_mnuInvHojaOpc_1"
        Me._mnuInvHojaOpc_1.Size = New System.Drawing.Size(184, 22)
        Me._mnuInvHojaOpc_1.Text = "Análisis comparativo"
        '
        'mnuInvOpc
        '
        '
        '_mnuInvOpc_0
        '
        Me._mnuInvOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuInvOpc_0.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuInvEntradasOpc_0, Me._mnuInvEntradasOpc_1, Me._mnuInvEntradasOpc_2, Me._mnuInvEntradasOpc_3, Me._mnuInvEntradasOpc_4, Me._mnuInvEntradasOpc_5, Me._mnuInvEntradasOpc_6, Me._mnuInvEntradasOpc_7})
        Me.mnuInvOpc.SetIndex(Me._mnuInvOpc_0, CType(0, Short))
        Me._mnuInvOpc_0.Name = "_mnuInvOpc_0"
        Me._mnuInvOpc_0.Size = New System.Drawing.Size(195, 22)
        Me._mnuInvOpc_0.Text = "Entradas"
        '
        '_mnuInvOpc_1
        '
        Me._mnuInvOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuInvOpc_1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuInvSalidasOpc_0, Me._mnuInvSalidasOpc_1, Me._mnuInvSalidasOpc_2, Me._mnuInvSalidasOpc_3, Me._mnuInvSalidasOpc_4, Me._mnuInvSalidasOpc_5, Me._mnuInvSalidasOpc_6, Me._mnuInvSalidasOpc_7})
        Me.mnuInvOpc.SetIndex(Me._mnuInvOpc_1, CType(1, Short))
        Me._mnuInvOpc_1.Name = "_mnuInvOpc_1"
        Me._mnuInvOpc_1.Size = New System.Drawing.Size(195, 22)
        Me._mnuInvOpc_1.Text = "Salidas"
        '
        '_mnuInvSalidasOpc_0
        '
        Me._mnuInvSalidasOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvSalidasOpc.SetIndex(Me._mnuInvSalidasOpc_0, CType(0, Short))
        Me._mnuInvSalidasOpc_0.Name = "_mnuInvSalidasOpc_0"
        Me._mnuInvSalidasOpc_0.Size = New System.Drawing.Size(244, 22)
        Me._mnuInvSalidasOpc_0.Text = "Por Venta"
        '
        '_mnuInvSalidasOpc_1
        '
        Me._mnuInvSalidasOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvSalidasOpc.SetIndex(Me._mnuInvSalidasOpc_1, CType(1, Short))
        Me._mnuInvSalidasOpc_1.Name = "_mnuInvSalidasOpc_1"
        Me._mnuInvSalidasOpc_1.Size = New System.Drawing.Size(244, 22)
        Me._mnuInvSalidasOpc_1.Text = "Por Transferencias"
        '
        '_mnuInvSalidasOpc_2
        '
        Me._mnuInvSalidasOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvSalidasOpc.SetIndex(Me._mnuInvSalidasOpc_2, CType(2, Short))
        Me._mnuInvSalidasOpc_2.Name = "_mnuInvSalidasOpc_2"
        Me._mnuInvSalidasOpc_2.Size = New System.Drawing.Size(244, 22)
        Me._mnuInvSalidasOpc_2.Text = "Por Devolución sobre Compra"
        '
        '_mnuInvSalidasOpc_3
        '
        Me._mnuInvSalidasOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvSalidasOpc.SetIndex(Me._mnuInvSalidasOpc_3, CType(3, Short))
        Me._mnuInvSalidasOpc_3.Name = "_mnuInvSalidasOpc_3"
        Me._mnuInvSalidasOpc_3.Size = New System.Drawing.Size(244, 22)
        Me._mnuInvSalidasOpc_3.Text = "A Vendedores Externos"
        '
        '_mnuInvSalidasOpc_4
        '
        Me._mnuInvSalidasOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvSalidasOpc.SetIndex(Me._mnuInvSalidasOpc_4, CType(4, Short))
        Me._mnuInvSalidasOpc_4.Name = "_mnuInvSalidasOpc_4"
        Me._mnuInvSalidasOpc_4.Size = New System.Drawing.Size(244, 22)
        Me._mnuInvSalidasOpc_4.Text = "Por Venta a Vendedores Externos"
        '
        '_mnuInvSalidasOpc_5
        '
        Me._mnuInvSalidasOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvSalidasOpc.SetIndex(Me._mnuInvSalidasOpc_5, CType(5, Short))
        Me._mnuInvSalidasOpc_5.Name = "_mnuInvSalidasOpc_5"
        Me._mnuInvSalidasOpc_5.Size = New System.Drawing.Size(244, 22)
        Me._mnuInvSalidasOpc_5.Text = "Por Obsequio"
        '
        '_mnuInvSalidasOpc_6
        '
        Me._mnuInvSalidasOpc_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvSalidasOpc.SetIndex(Me._mnuInvSalidasOpc_6, CType(6, Short))
        Me._mnuInvSalidasOpc_6.Name = "_mnuInvSalidasOpc_6"
        Me._mnuInvSalidasOpc_6.Size = New System.Drawing.Size(244, 22)
        Me._mnuInvSalidasOpc_6.Text = "Por Préstamos de Artículo"
        '
        '_mnuInvSalidasOpc_7
        '
        Me._mnuInvSalidasOpc_7.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvSalidasOpc.SetIndex(Me._mnuInvSalidasOpc_7, CType(7, Short))
        Me._mnuInvSalidasOpc_7.Name = "_mnuInvSalidasOpc_7"
        Me._mnuInvSalidasOpc_7.Size = New System.Drawing.Size(244, 22)
        Me._mnuInvSalidasOpc_7.Text = "Por Sustitución"
        '
        '_mnuInvOpc_2
        '
        Me._mnuInvOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvOpc.SetIndex(Me._mnuInvOpc_2, CType(2, Short))
        Me._mnuInvOpc_2.Name = "_mnuInvOpc_2"
        Me._mnuInvOpc_2.Size = New System.Drawing.Size(195, 22)
        Me._mnuInvOpc_2.Text = "Impresión de Etiquetas"
        '
        '_mnuInvOpc_3
        '
        Me._mnuInvOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvOpc.SetIndex(Me._mnuInvOpc_3, CType(3, Short))
        Me._mnuInvOpc_3.Name = "_mnuInvOpc_3"
        Me._mnuInvOpc_3.Size = New System.Drawing.Size(195, 22)
        Me._mnuInvOpc_3.Text = "Stock Básico de Tienda"
        '
        '_mnuInvOpc_4
        '
        Me._mnuInvOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuInvOpc_4.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuInvHojaOpc_0, Me._mnuInvHojaOpc_1})
        Me.mnuInvOpc.SetIndex(Me._mnuInvOpc_4, CType(4, Short))
        Me._mnuInvOpc_4.Name = "_mnuInvOpc_4"
        Me._mnuInvOpc_4.Size = New System.Drawing.Size(195, 22)
        Me._mnuInvOpc_4.Text = "Inventario físico"
        '
        '_mnuInvOpc_5
        '
        Me._mnuInvOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuInvOpc_5.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuInvRptOpc_0, Me._mnuInvRptOpc_1, Me._mnuInvRptOpc_2, Me._mnuInvRptOpc_3, Me._mnuInvRptOpc_4})
        Me.mnuInvOpc.SetIndex(Me._mnuInvOpc_5, CType(5, Short))
        Me._mnuInvOpc_5.Name = "_mnuInvOpc_5"
        Me._mnuInvOpc_5.Size = New System.Drawing.Size(195, 22)
        Me._mnuInvOpc_5.Text = "Reportes"
        '
        '_mnuInvRptOpc_0
        '
        Me._mnuInvRptOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvRptOpc.SetIndex(Me._mnuInvRptOpc_0, CType(0, Short))
        Me._mnuInvRptOpc_0.Name = "_mnuInvRptOpc_0"
        Me._mnuInvRptOpc_0.Size = New System.Drawing.Size(253, 22)
        Me._mnuInvRptOpc_0.Text = "Kardex"
        '
        '_mnuInvRptOpc_1
        '
        Me._mnuInvRptOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvRptOpc.SetIndex(Me._mnuInvRptOpc_1, CType(1, Short))
        Me._mnuInvRptOpc_1.Name = "_mnuInvRptOpc_1"
        Me._mnuInvRptOpc_1.Size = New System.Drawing.Size(253, 22)
        Me._mnuInvRptOpc_1.Text = "Existencias y Costos"
        '
        '_mnuInvRptOpc_2
        '
        Me._mnuInvRptOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvRptOpc.SetIndex(Me._mnuInvRptOpc_2, CType(2, Short))
        Me._mnuInvRptOpc_2.Name = "_mnuInvRptOpc_2"
        Me._mnuInvRptOpc_2.Size = New System.Drawing.Size(253, 22)
        Me._mnuInvRptOpc_2.Text = "Reporte de préstamos de artículos"
        '
        '_mnuInvRptOpc_3
        '
        Me._mnuInvRptOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvRptOpc.SetIndex(Me._mnuInvRptOpc_3, CType(3, Short))
        Me._mnuInvRptOpc_3.Name = "_mnuInvRptOpc_3"
        Me._mnuInvRptOpc_3.Size = New System.Drawing.Size(253, 22)
        Me._mnuInvRptOpc_3.Text = "Comparación Existencia-Stock"
        '
        '_mnuInvRptOpc_4
        '
        Me._mnuInvRptOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuInvRptOpc.SetIndex(Me._mnuInvRptOpc_4, CType(4, Short))
        Me._mnuInvRptOpc_4.Name = "_mnuInvRptOpc_4"
        Me._mnuInvRptOpc_4.Size = New System.Drawing.Size(253, 22)
        Me._mnuInvRptOpc_4.Text = "Transferencias no conciliadas"
        '
        'mnuInvRptOpc
        '
        '
        'mnuInvSalidasOpc
        '
        '
        'mnuSegOpc
        '
        '
        '_mnuSegOpc_0
        '
        Me._mnuSegOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuSegOpc.SetIndex(Me._mnuSegOpc_0, CType(0, Short))
        Me._mnuSegOpc_0.Name = "_mnuSegOpc_0"
        Me._mnuSegOpc_0.Size = New System.Drawing.Size(229, 22)
        Me._mnuSegOpc_0.Text = "&Rastreo Módulos y Funciones"
        '
        '_mnuSegOpc_1
        '
        Me._mnuSegOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuSegOpc.SetIndex(Me._mnuSegOpc_1, CType(1, Short))
        Me._mnuSegOpc_1.Name = "_mnuSegOpc_1"
        Me._mnuSegOpc_1.Size = New System.Drawing.Size(229, 22)
        Me._mnuSegOpc_1.Text = "&Módulos y Funciones"
        '
        '_mnuSegOpc_2
        '
        Me._mnuSegOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuSegOpc.SetIndex(Me._mnuSegOpc_2, CType(2, Short))
        Me._mnuSegOpc_2.Name = "_mnuSegOpc_2"
        Me._mnuSegOpc_2.Size = New System.Drawing.Size(229, 22)
        Me._mnuSegOpc_2.Text = "&Usuarios y accesos"
        '
        'mnuVentasOpc
        '
        '
        '_mnuVentasOpc_0
        '
        Me._mnuVentasOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuVentasOpc_0.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuVentasSalMerOpc_0, Me._mnuVentasSalMerOpc_1, Me._mnuVentasSalMerOpc_2, Me._mnuVentasSalMerOpc_3, Me._mnuVentasSalMerOpc_4, Me._mnuVentasSalMerOpc_5, Me._mnuVentasSalMerOpc_6, Me._mnuVentasSalMerOpc_7, Me._mnuVentasSalMerOpc_8, Me._mnuVentasSalMerOpc_9, Me._mnuVentasSalMerOpc_10, Me._mnuVentasSalMerOpc_11, Me._mnuVentasSalMerOpc_12, Me._mnuVentasSalMerOpc_13})
        Me.mnuVentasOpc.SetIndex(Me._mnuVentasOpc_0, CType(0, Short))
        Me._mnuVentasOpc_0.Name = "_mnuVentasOpc_0"
        Me._mnuVentasOpc_0.Size = New System.Drawing.Size(244, 22)
        Me._mnuVentasOpc_0.Text = "Ventas Salida de Mercancía"
        '
        '_mnuVentasSalMerOpc_0
        '
        Me._mnuVentasSalMerOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_0, CType(0, Short))
        Me._mnuVentasSalMerOpc_0.Name = "_mnuVentasSalMerOpc_0"
        Me._mnuVentasSalMerOpc_0.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_0.Text = "Por Período y Tienda"
        '
        '_mnuVentasSalMerOpc_1
        '
        Me._mnuVentasSalMerOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_1, CType(1, Short))
        Me._mnuVentasSalMerOpc_1.Name = "_mnuVentasSalMerOpc_1"
        Me._mnuVentasSalMerOpc_1.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_1.Text = "Por Proveedor y Tienda"
        '
        '_mnuVentasSalMerOpc_2
        '
        Me._mnuVentasSalMerOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_2, CType(2, Short))
        Me._mnuVentasSalMerOpc_2.Name = "_mnuVentasSalMerOpc_2"
        Me._mnuVentasSalMerOpc_2.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_2.Text = "Por Clasificación"
        '
        '_mnuVentasSalMerOpc_3
        '
        Me._mnuVentasSalMerOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_3, CType(3, Short))
        Me._mnuVentasSalMerOpc_3.Name = "_mnuVentasSalMerOpc_3"
        Me._mnuVentasSalMerOpc_3.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_3.Text = "Comparativo de Ventas Diarias con año Anterior (en Excel)"
        '
        '_mnuVentasSalMerOpc_4
        '
        Me._mnuVentasSalMerOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuVentasSalMerOpc_4.Enabled = False
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_4, CType(4, Short))
        Me._mnuVentasSalMerOpc_4.Name = "_mnuVentasSalMerOpc_4"
        Me._mnuVentasSalMerOpc_4.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_4.Text = "Utilidad por Línea"
        Me._mnuVentasSalMerOpc_4.Visible = False
        '
        '_mnuVentasSalMerOpc_5
        '
        Me._mnuVentasSalMerOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_5, CType(5, Short))
        Me._mnuVentasSalMerOpc_5.Name = "_mnuVentasSalMerOpc_5"
        Me._mnuVentasSalMerOpc_5.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_5.Text = "Relojería por Marca Modelo"
        '
        '_mnuVentasSalMerOpc_6
        '
        Me._mnuVentasSalMerOpc_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_6, CType(6, Short))
        Me._mnuVentasSalMerOpc_6.Name = "_mnuVentasSalMerOpc_6"
        Me._mnuVentasSalMerOpc_6.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_6.Text = "Relojería por Material de Fabricación"
        '
        '_mnuVentasSalMerOpc_7
        '
        Me._mnuVentasSalMerOpc_7.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_7, CType(7, Short))
        Me._mnuVentasSalMerOpc_7.Name = "_mnuVentasSalMerOpc_7"
        Me._mnuVentasSalMerOpc_7.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_7.Text = "Flujo de Venta por Proveedor"
        '
        '_mnuVentasSalMerOpc_8
        '
        Me._mnuVentasSalMerOpc_8.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_8, CType(8, Short))
        Me._mnuVentasSalMerOpc_8.Name = "_mnuVentasSalMerOpc_8"
        Me._mnuVentasSalMerOpc_8.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_8.Text = "Ventas Por Cliente"
        '
        '_mnuVentasSalMerOpc_9
        '
        Me._mnuVentasSalMerOpc_9.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_9, CType(9, Short))
        Me._mnuVentasSalMerOpc_9.Name = "_mnuVentasSalMerOpc_9"
        Me._mnuVentasSalMerOpc_9.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_9.Text = "Ventas por Vendedor"
        '
        '_mnuVentasSalMerOpc_10
        '
        Me._mnuVentasSalMerOpc_10.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_10, CType(10, Short))
        Me._mnuVentasSalMerOpc_10.Name = "_mnuVentasSalMerOpc_10"
        Me._mnuVentasSalMerOpc_10.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_10.Text = "Comisiones por Vendedor"
        '
        '_mnuVentasSalMerOpc_11
        '
        Me._mnuVentasSalMerOpc_11.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_11, CType(11, Short))
        Me._mnuVentasSalMerOpc_11.Name = "_mnuVentasSalMerOpc_11"
        Me._mnuVentasSalMerOpc_11.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_11.Text = "Listado de Ventas"
        '
        '_mnuVentasSalMerOpc_12
        '
        Me._mnuVentasSalMerOpc_12.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuVentasSalMerOpc_12.Name = "_mnuVentasSalMerOpc_12"
        Me._mnuVentasSalMerOpc_12.Size = New System.Drawing.Size(376, 6)
        '
        '_mnuVentasSalMerOpc_13
        '
        Me._mnuVentasSalMerOpc_13.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuVentasSalMerOpc_13.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuVentasSalMerOpcRepEjec_0, Me._mnuVentasSalMerOpcRepEjec_1, Me._mnuVentasSalMerOpcRepEjec_2, Me._mnuVentasSalMerOpcRepEjec_3, Me._mnuVentasSalMerOpcRepEjec_4, Me._mnuVentasSalMerOpcRepEjec_5})
        Me.mnuVentasSalMerOpc.SetIndex(Me._mnuVentasSalMerOpc_13, CType(13, Short))
        Me._mnuVentasSalMerOpc_13.Name = "_mnuVentasSalMerOpc_13"
        Me._mnuVentasSalMerOpc_13.Size = New System.Drawing.Size(379, 22)
        Me._mnuVentasSalMerOpc_13.Text = "Reportes Ejecutivos"
        '
        '_mnuVentasSalMerOpcRepEjec_0
        '
        Me._mnuVentasSalMerOpcRepEjec_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpcRepEjec.SetIndex(Me._mnuVentasSalMerOpcRepEjec_0, CType(0, Short))
        Me._mnuVentasSalMerOpcRepEjec_0.Name = "_mnuVentasSalMerOpcRepEjec_0"
        Me._mnuVentasSalMerOpcRepEjec_0.Size = New System.Drawing.Size(333, 22)
        Me._mnuVentasSalMerOpcRepEjec_0.Text = "Reporte de Ventas y Existencias por Proveedor"
        '
        '_mnuVentasSalMerOpcRepEjec_1
        '
        Me._mnuVentasSalMerOpcRepEjec_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpcRepEjec.SetIndex(Me._mnuVentasSalMerOpcRepEjec_1, CType(1, Short))
        Me._mnuVentasSalMerOpcRepEjec_1.Name = "_mnuVentasSalMerOpcRepEjec_1"
        Me._mnuVentasSalMerOpcRepEjec_1.Size = New System.Drawing.Size(333, 22)
        Me._mnuVentasSalMerOpcRepEjec_1.Text = "Reporte de Ventas por Resurtir"
        '
        '_mnuVentasSalMerOpcRepEjec_2
        '
        Me._mnuVentasSalMerOpcRepEjec_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpcRepEjec.SetIndex(Me._mnuVentasSalMerOpcRepEjec_2, CType(2, Short))
        Me._mnuVentasSalMerOpcRepEjec_2.Name = "_mnuVentasSalMerOpcRepEjec_2"
        Me._mnuVentasSalMerOpcRepEjec_2.Size = New System.Drawing.Size(333, 22)
        Me._mnuVentasSalMerOpcRepEjec_2.Text = "Reporte de Ventas y Utilidad Global por Grupo"
        '
        '_mnuVentasSalMerOpcRepEjec_3
        '
        Me._mnuVentasSalMerOpcRepEjec_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpcRepEjec.SetIndex(Me._mnuVentasSalMerOpcRepEjec_3, CType(3, Short))
        Me._mnuVentasSalMerOpcRepEjec_3.Name = "_mnuVentasSalMerOpcRepEjec_3"
        Me._mnuVentasSalMerOpcRepEjec_3.Size = New System.Drawing.Size(333, 22)
        Me._mnuVentasSalMerOpcRepEjec_3.Text = "Reporte de Ventas Salida de Mercancia por Grupo"
        '
        '_mnuVentasSalMerOpcRepEjec_4
        '
        Me._mnuVentasSalMerOpcRepEjec_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpcRepEjec.SetIndex(Me._mnuVentasSalMerOpcRepEjec_4, CType(4, Short))
        Me._mnuVentasSalMerOpcRepEjec_4.Name = "_mnuVentasSalMerOpcRepEjec_4"
        Me._mnuVentasSalMerOpcRepEjec_4.Size = New System.Drawing.Size(333, 22)
        Me._mnuVentasSalMerOpcRepEjec_4.Text = "Reporte de Utilidad por Grupo"
        '
        '_mnuVentasSalMerOpcRepEjec_5
        '
        Me._mnuVentasSalMerOpcRepEjec_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasSalMerOpcRepEjec.SetIndex(Me._mnuVentasSalMerOpcRepEjec_5, CType(5, Short))
        Me._mnuVentasSalMerOpcRepEjec_5.Name = "_mnuVentasSalMerOpcRepEjec_5"
        Me._mnuVentasSalMerOpcRepEjec_5.Size = New System.Drawing.Size(333, 22)
        Me._mnuVentasSalMerOpcRepEjec_5.Text = "Reporte de Ventas y Existencias por Familia"
        '
        '_mnuVentasOpc_1
        '
        Me._mnuVentasOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuVentasOpc_1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuVentasVtasIngrOpc_0, Me._mnuVentasVtasIngrOpc_1, Me._mnuVentasVtasIngrOpc_2, Me._mnuVentasVtasIngrOpc_3, Me._mnuVentasVtasIngrOpc_4})
        Me.mnuVentasOpc.SetIndex(Me._mnuVentasOpc_1, CType(1, Short))
        Me._mnuVentasOpc_1.Name = "_mnuVentasOpc_1"
        Me._mnuVentasOpc_1.Size = New System.Drawing.Size(244, 22)
        Me._mnuVentasOpc_1.Text = "Ventas Ingresos"
        '
        '_mnuVentasVtasIngrOpc_0
        '
        Me._mnuVentasVtasIngrOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasVtasIngrOpc.SetIndex(Me._mnuVentasVtasIngrOpc_0, CType(0, Short))
        Me._mnuVentasVtasIngrOpc_0.Name = "_mnuVentasVtasIngrOpc_0"
        Me._mnuVentasVtasIngrOpc_0.Size = New System.Drawing.Size(240, 22)
        Me._mnuVentasVtasIngrOpc_0.Text = "Ingresos Generales"
        '
        '_mnuVentasVtasIngrOpc_1
        '
        Me._mnuVentasVtasIngrOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasVtasIngrOpc.SetIndex(Me._mnuVentasVtasIngrOpc_1, CType(1, Short))
        Me._mnuVentasVtasIngrOpc_1.Name = "_mnuVentasVtasIngrOpc_1"
        Me._mnuVentasVtasIngrOpc_1.Size = New System.Drawing.Size(240, 22)
        Me._mnuVentasVtasIngrOpc_1.Text = "Por Periodo y Tienda"
        '
        '_mnuVentasVtasIngrOpc_2
        '
        Me._mnuVentasVtasIngrOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasVtasIngrOpc.SetIndex(Me._mnuVentasVtasIngrOpc_2, CType(2, Short))
        Me._mnuVentasVtasIngrOpc_2.Name = "_mnuVentasVtasIngrOpc_2"
        Me._mnuVentasVtasIngrOpc_2.Size = New System.Drawing.Size(240, 22)
        Me._mnuVentasVtasIngrOpc_2.Text = "Abonos"
        '
        '_mnuVentasVtasIngrOpc_3
        '
        Me._mnuVentasVtasIngrOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasVtasIngrOpc.SetIndex(Me._mnuVentasVtasIngrOpc_3, CType(3, Short))
        Me._mnuVentasVtasIngrOpc_3.Name = "_mnuVentasVtasIngrOpc_3"
        Me._mnuVentasVtasIngrOpc_3.Size = New System.Drawing.Size(240, 22)
        Me._mnuVentasVtasIngrOpc_3.Text = "Por Reparación"
        '
        '_mnuVentasVtasIngrOpc_4
        '
        Me._mnuVentasVtasIngrOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasVtasIngrOpc.SetIndex(Me._mnuVentasVtasIngrOpc_4, CType(4, Short))
        Me._mnuVentasVtasIngrOpc_4.Name = "_mnuVentasVtasIngrOpc_4"
        Me._mnuVentasVtasIngrOpc_4.Size = New System.Drawing.Size(240, 22)
        Me._mnuVentasVtasIngrOpc_4.Text = "Ingresos por Concepto de Pago"
        '
        '_mnuVentasOpc_2
        '
        Me._mnuVentasOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuVentasOpc_2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuVentasVendExtOpc_0, Me._mnuVentasVendExtOpc_1, Me._mnuVentasVendExtOpc_2, Me._mnuVentasVendExtOpc_3, Me._mnuVentasVendExtOpc_4, Me._mnuVentasVendExtOpc_5})
        Me.mnuVentasOpc.SetIndex(Me._mnuVentasOpc_2, CType(2, Short))
        Me._mnuVentasOpc_2.Name = "_mnuVentasOpc_2"
        Me._mnuVentasOpc_2.Size = New System.Drawing.Size(244, 22)
        Me._mnuVentasOpc_2.Text = "Vendedor Externo"
        '
        '_mnuVentasVendExtOpc_0
        '
        Me._mnuVentasVendExtOpc_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasVendExtOpc.SetIndex(Me._mnuVentasVendExtOpc_0, CType(0, Short))
        Me._mnuVentasVendExtOpc_0.Name = "_mnuVentasVendExtOpc_0"
        Me._mnuVentasVendExtOpc_0.Size = New System.Drawing.Size(287, 22)
        Me._mnuVentasVendExtOpc_0.Text = "Recepción de Mercancia"
        '
        '_mnuVentasVendExtOpc_1
        '
        Me._mnuVentasVendExtOpc_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasVendExtOpc.SetIndex(Me._mnuVentasVendExtOpc_1, CType(1, Short))
        Me._mnuVentasVendExtOpc_1.Name = "_mnuVentasVendExtOpc_1"
        Me._mnuVentasVendExtOpc_1.Size = New System.Drawing.Size(287, 22)
        Me._mnuVentasVendExtOpc_1.Text = "Entrega de Mercancía"
        '
        '_mnuVentasVendExtOpc_2
        '
        Me._mnuVentasVendExtOpc_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasVendExtOpc.SetIndex(Me._mnuVentasVendExtOpc_2, CType(2, Short))
        Me._mnuVentasVendExtOpc_2.Name = "_mnuVentasVendExtOpc_2"
        Me._mnuVentasVendExtOpc_2.Size = New System.Drawing.Size(287, 22)
        Me._mnuVentasVendExtOpc_2.Text = "Reporte de Existencias"
        '
        '_mnuVentasVendExtOpc_3
        '
        Me._mnuVentasVendExtOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasVendExtOpc.SetIndex(Me._mnuVentasVendExtOpc_3, CType(3, Short))
        Me._mnuVentasVendExtOpc_3.Name = "_mnuVentasVendExtOpc_3"
        Me._mnuVentasVendExtOpc_3.Size = New System.Drawing.Size(287, 22)
        Me._mnuVentasVendExtOpc_3.Text = "Liquidación"
        '
        '_mnuVentasVendExtOpc_4
        '
        Me._mnuVentasVendExtOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasVendExtOpc.SetIndex(Me._mnuVentasVendExtOpc_4, CType(4, Short))
        Me._mnuVentasVendExtOpc_4.Name = "_mnuVentasVendExtOpc_4"
        Me._mnuVentasVendExtOpc_4.Size = New System.Drawing.Size(287, 22)
        Me._mnuVentasVendExtOpc_4.Text = "Ingresos por Entrega de Mercancía"
        '
        '_mnuVentasVendExtOpc_5
        '
        Me._mnuVentasVendExtOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasVendExtOpc.SetIndex(Me._mnuVentasVendExtOpc_5, CType(5, Short))
        Me._mnuVentasVendExtOpc_5.Name = "_mnuVentasVendExtOpc_5"
        Me._mnuVentasVendExtOpc_5.Size = New System.Drawing.Size(287, 22)
        Me._mnuVentasVendExtOpc_5.Text = "Reporte Detallado de Recepción/Entrega"
        '
        '_mnuVentasOpc_3
        '
        Me._mnuVentasOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasOpc.SetIndex(Me._mnuVentasOpc_3, CType(3, Short))
        Me._mnuVentasOpc_3.Name = "_mnuVentasOpc_3"
        Me._mnuVentasOpc_3.Size = New System.Drawing.Size(244, 22)
        Me._mnuVentasOpc_3.Text = "Apartados"
        '
        '_mnuVentasOpc_4
        '
        Me._mnuVentasOpc_4.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasOpc.SetIndex(Me._mnuVentasOpc_4, CType(4, Short))
        Me._mnuVentasOpc_4.Name = "_mnuVentasOpc_4"
        Me._mnuVentasOpc_4.Size = New System.Drawing.Size(244, 22)
        Me._mnuVentasOpc_4.Text = "Reparaciones"
        '
        '_mnuVentasOpc_5
        '
        Me._mnuVentasOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasOpc.SetIndex(Me._mnuVentasOpc_5, CType(5, Short))
        Me._mnuVentasOpc_5.Name = "_mnuVentasOpc_5"
        Me._mnuVentasOpc_5.Size = New System.Drawing.Size(244, 22)
        Me._mnuVentasOpc_5.Text = "Cuentas por Cobrar"
        '
        '_mnuVentasOpc_7
        '
        Me._mnuVentasOpc_7.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasOpc.SetIndex(Me._mnuVentasOpc_7, CType(7, Short))
        Me._mnuVentasOpc_7.Name = "_mnuVentasOpc_7"
        Me._mnuVentasOpc_7.Size = New System.Drawing.Size(244, 22)
        Me._mnuVentasOpc_7.Text = "Estado de Resultados (En Excel)"
        '
        '_mnuVentasOpc_8
        '
        Me._mnuVentasOpc_8.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasOpc.SetIndex(Me._mnuVentasOpc_8, CType(8, Short))
        Me._mnuVentasOpc_8.Name = "_mnuVentasOpc_8"
        Me._mnuVentasOpc_8.Size = New System.Drawing.Size(244, 22)
        Me._mnuVentasOpc_8.Text = "Relación de Gastos"
        '
        '_mnuVentasOpc_10
        '
        Me._mnuVentasOpc_10.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasOpc.SetIndex(Me._mnuVentasOpc_10, CType(10, Short))
        Me._mnuVentasOpc_10.Name = "_mnuVentasOpc_10"
        Me._mnuVentasOpc_10.Size = New System.Drawing.Size(244, 22)
        Me._mnuVentasOpc_10.Text = "Administración de Reparaciones"
        '
        '_mnuVentasOpc_11
        '
        Me._mnuVentasOpc_11.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.mnuVentasOpc.SetIndex(Me._mnuVentasOpc_11, CType(11, Short))
        Me._mnuVentasOpc_11.Name = "_mnuVentasOpc_11"
        Me._mnuVentasOpc_11.Size = New System.Drawing.Size(244, 22)
        Me._mnuVentasOpc_11.Text = "Verificador de Precios"
        '
        'mnuVentasSalMerOpc
        '
        '
        'mnuVentasSalMerOpcRepEjec
        '
        '
        'mnuVentasVendExtOpc
        '
        '
        'mnuVentasVtasIngrOpc
        '
        '
        'MainMenu1
        '
        Me.MainMenu1.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.MainMenu1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCatalogos, Me.mnuVentas, Me.mnuComprasyCxP, Me.mnuFacturacion, Me.mnuBancos, Me.mnuInventarios, Me.mnuConfiguracion, Me.mnuSeg, Me.menuContextualGen, Me._MenuAcercaDe_0})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(1043, 24)
        Me.MainMenu1.TabIndex = 2
        '
        'mnuCatalogos
        '
        Me.mnuCatalogos.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.mnuCatalogos.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuCatalogosOpc_0, Me._mnuCatalogosOpc_1, Me._mnuCatalogosOpc_2, Me._mnuCatalogosOpc_3, Me._mnuCatalogosOpc_4, Me._mnuCatalogosOpc_5, Me._mnuCatalogosOpc_6, Me._mnuCatalogosOpc_7, Me._mnuCatalogosOpc_8, Me._mnuCatalogosOpc_9, Me._mnuCatalogosOpc_10, Me._mnuCatalogosOpc_11, Me._mnuCatalogosOpc_12, Me._mnuCatalogosOpc_13, Me._mnuCatalogosOpc_14, Me._mnuCatalogosOpc_15, Me._mnuCatalogosOpc_16, Me._mnuCatalogosOpc_17, Me._mnuCatalogosOpc_18, Me._mnuCatalogosOpc_19, Me._mnuCatalogosOpc_20, Me._mnuCatalogosOpc_21, Me._mnuCatalogosOpc_22, Me._mnuCatalogosOpc_23, Me._mnuCatalogosOpc_24})
        Me.mnuCatalogos.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuCatalogos.Name = "mnuCatalogos"
        Me.mnuCatalogos.Size = New System.Drawing.Size(72, 20)
        Me.mnuCatalogos.Text = "&Catálogos"
        '
        '_mnuCatalogosOpc_15
        '
        Me._mnuCatalogosOpc_15.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuCatalogosOpc_15.Name = "_mnuCatalogosOpc_15"
        Me._mnuCatalogosOpc_15.Size = New System.Drawing.Size(274, 6)
        '
        '_mnuCatalogosOpc_18
        '
        Me._mnuCatalogosOpc_18.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuCatalogosOpc_18.Name = "_mnuCatalogosOpc_18"
        Me._mnuCatalogosOpc_18.Size = New System.Drawing.Size(274, 6)
        '
        '_mnuCatalogosOpc_21
        '
        Me._mnuCatalogosOpc_21.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuCatalogosOpc_21.Name = "_mnuCatalogosOpc_21"
        Me._mnuCatalogosOpc_21.Size = New System.Drawing.Size(274, 6)
        '
        'mnuVentas
        '
        Me.mnuVentas.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.mnuVentas.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuVentasOpc_0, Me._mnuVentasOpc_1, Me._mnuVentasOpc_2, Me._mnuVentasOpc_3, Me._mnuVentasOpc_4, Me._mnuVentasOpc_5, Me._mnuVentasOpc_6, Me._mnuVentasOpc_7, Me._mnuVentasOpc_8, Me._mnuVentasOpc_9, Me._mnuVentasOpc_10, Me._mnuVentasOpc_11})
        Me.mnuVentas.ForeColor = System.Drawing.Color.Black
        Me.mnuVentas.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuVentas.Name = "mnuVentas"
        Me.mnuVentas.Size = New System.Drawing.Size(53, 20)
        Me.mnuVentas.Text = "Ven&tas"
        '
        '_mnuVentasOpc_6
        '
        Me._mnuVentasOpc_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuVentasOpc_6.ForeColor = System.Drawing.Color.Black
        Me._mnuVentasOpc_6.Name = "_mnuVentasOpc_6"
        Me._mnuVentasOpc_6.Size = New System.Drawing.Size(241, 6)
        '
        '_mnuVentasOpc_9
        '
        Me._mnuVentasOpc_9.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuVentasOpc_9.ForeColor = System.Drawing.Color.Black
        Me._mnuVentasOpc_9.Name = "_mnuVentasOpc_9"
        Me._mnuVentasOpc_9.Size = New System.Drawing.Size(241, 6)
        '
        'mnuComprasyCxP
        '
        Me.mnuComprasyCxP.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.mnuComprasyCxP.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuCompyCxPOpc_0, Me._mnuCompyCxPOpc_1, Me._mnuCompyCxPOpc_2, Me._mnuCompyCxPOpc_3, Me._mnuCompyCxPOpc_4, Me._mnuCompyCxPOpc_5, Me._mnuCompyCxPOpc_6, Me._mnuCompyCxPOpc_7, Me._mnuCompyCxPOpc_8, Me._mnuCompyCxPOpc_9})
        Me.mnuComprasyCxP.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuComprasyCxP.Name = "mnuComprasyCxP"
        Me.mnuComprasyCxP.Size = New System.Drawing.Size(99, 20)
        Me.mnuComprasyCxP.Text = "Compras y C&xP"
        '
        '_mnuCompyCxPOpc_8
        '
        Me._mnuCompyCxPOpc_8.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuCompyCxPOpc_8.Name = "_mnuCompyCxPOpc_8"
        Me._mnuCompyCxPOpc_8.Size = New System.Drawing.Size(244, 6)
        '
        'mnuFacturacion
        '
        Me.mnuFacturacion.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.mnuFacturacion.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuFacturacionOpc_0, Me._mnuFacturacionOpc_1, Me._mnuFacturacionOpc_2, Me._mnuFacturacionOpc_3, Me._mnuFacturacionOpc_4, Me._mnuFacturacionOpc_5})
        Me.mnuFacturacion.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuFacturacion.Name = "mnuFacturacion"
        Me.mnuFacturacion.Size = New System.Drawing.Size(81, 20)
        Me.mnuFacturacion.Text = "&Facturación"
        '
        '_mnuFacturacionOpc_3
        '
        Me._mnuFacturacionOpc_3.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuFacturacionOpc_3.Name = "_mnuFacturacionOpc_3"
        Me._mnuFacturacionOpc_3.Size = New System.Drawing.Size(232, 6)
        '
        'mnuBancos
        '
        Me.mnuBancos.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.mnuBancos.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuBancosOpc_0, Me._mnuBancosOpc_1, Me._mnuBancosOpc_2, Me._mnuBancosOpc_3})
        Me.mnuBancos.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuBancos.Name = "mnuBancos"
        Me.mnuBancos.Size = New System.Drawing.Size(57, 20)
        Me.mnuBancos.Text = "&Bancos"
        '
        'mnuInventarios
        '
        Me.mnuInventarios.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.mnuInventarios.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuInvOpc_0, Me._mnuInvOpc_1, Me._mnuInvOpc_2, Me._mnuInvOpc_3, Me._mnuInvOpc_4, Me._mnuInvOpc_5})
        Me.mnuInventarios.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuInventarios.Name = "mnuInventarios"
        Me.mnuInventarios.Size = New System.Drawing.Size(77, 20)
        Me.mnuInventarios.Text = "&Inventarios"
        '
        'mnuConfiguracion
        '
        Me.mnuConfiguracion.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.mnuConfiguracion.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuConfiguracionOpc_0, Me._mnuConfiguracionOpc_1, Me._mnuConfiguracionOpc_2, Me._mnuConfiguracionOpc_3, Me._mnuConfiguracionOpc_4, Me._mnuConfiguracionOpc_5, Me._mnuConfiguracionOpc_6, Me._mnuConfiguracionOpc_7, Me._mnuConfiguracionOpc_8, Me._mnuConfiguracionOpc_9, Me._mnuConfiguracionOpc_10, Me._mnuConfiguracionOpc_11, Me._mnuConfiguracionOpc_12})
        Me.mnuConfiguracion.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuConfiguracion.Name = "mnuConfiguracion"
        Me.mnuConfiguracion.Size = New System.Drawing.Size(95, 20)
        Me.mnuConfiguracion.Text = "Con&figuración"
        '
        '_mnuConfiguracionOpc_6
        '
        Me._mnuConfiguracionOpc_6.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuConfiguracionOpc_6.Name = "_mnuConfiguracionOpc_6"
        Me._mnuConfiguracionOpc_6.Size = New System.Drawing.Size(299, 6)
        '
        '_mnuConfiguracionOpc_8
        '
        Me._mnuConfiguracionOpc_8.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuConfiguracionOpc_8.Name = "_mnuConfiguracionOpc_8"
        Me._mnuConfiguracionOpc_8.Size = New System.Drawing.Size(299, 6)
        '
        '_mnuConfiguracionOpc_11
        '
        Me._mnuConfiguracionOpc_11.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._mnuConfiguracionOpc_11.Name = "_mnuConfiguracionOpc_11"
        Me._mnuConfiguracionOpc_11.Size = New System.Drawing.Size(299, 6)
        '
        'mnuSeg
        '
        Me.mnuSeg.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.mnuSeg.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuSegOpc_0, Me._mnuSegOpc_1, Me._mnuSegOpc_2})
        Me.mnuSeg.ForeColor = System.Drawing.Color.Black
        Me.mnuSeg.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.mnuSeg.Name = "mnuSeg"
        Me.mnuSeg.Size = New System.Drawing.Size(72, 20)
        Me.mnuSeg.Text = "&Seguridad"
        '
        'menuContextualGen
        '
        Me.menuContextualGen.BackColor = System.Drawing.SystemColors.Highlight
        Me.menuContextualGen.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._menuContextualGenOpc_0, Me._menuContextualGenOpc_1, Me._menuContextualGenOpc_2, Me._menuContextualGenOpc_3, Me._menuContextualGenOpc_4, Me._menuContextualGenOpc_5, Me._menuContextualGenOpc_6, Me._menuContextualGenOpc_7, Me._menuContextualGenOpc_8, Me._menuContextualGenOpc_9, Me._menuContextualGenOpc_10})
        Me.menuContextualGen.MergeAction = System.Windows.Forms.MergeAction.Remove
        Me.menuContextualGen.Name = "menuContextualGen"
        Me.menuContextualGen.Size = New System.Drawing.Size(153, 20)
        Me.menuContextualGen.Text = "Menu Contextual General"
        Me.menuContextualGen.Visible = False
        '
        '_menuContextualGenOpc_5
        '
        Me._menuContextualGenOpc_5.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me._menuContextualGenOpc_5.Name = "_menuContextualGenOpc_5"
        Me._menuContextualGenOpc_5.Size = New System.Drawing.Size(122, 6)
        '
        'ButtonContainer
        '
        Me.ButtonContainer.AllowDrop = True
        Me.ButtonContainer.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(209, Byte), Integer), CType(CType(218, Byte), Integer))
        Me.ButtonContainer.Controls.Add(Me.btnSoporte)
        Me.ButtonContainer.Controls.Add(Me.ButtonTeleMarketing)
        Me.ButtonContainer.Controls.Add(Me.ButtonCorteDiario)
        Me.ButtonContainer.Controls.Add(Me.ButtonRegistroCobranza)
        Me.ButtonContainer.Controls.Add(Me.ButtonRegistroGastos)
        Me.ButtonContainer.Controls.Add(Me.ButtonConsultaInventario)
        Me.ButtonContainer.Controls.Add(Me.ButtonCompraEmergencia)
        Me.ButtonContainer.Controls.Add(Me.ButtonSalidasAOrden)
        Me.ButtonContainer.Controls.Add(Me.ButtonRecepcionProducto)
        Me.ButtonContainer.Controls.Add(Me.ButtonOrdenCompra)
        Me.ButtonContainer.Controls.Add(Me.ButtonHistorial)
        Me.ButtonContainer.Controls.Add(Me.ButtonCalendario)
        Me.ButtonContainer.Controls.Add(Me.ButtonCalculadora)
        Me.ButtonContainer.Controls.Add(Me.ButtonCotizacion)
        Me.ButtonContainer.Controls.Add(Me.ButtonEmpresas)
        Me.ButtonContainer.Controls.Add(Me.ButtonClientes)
        Me.ButtonContainer.Controls.Add(Me.ButtonRecepcion)
        Me.ButtonContainer.Controls.Add(Me.ButtonPanelControl)
        Me.ButtonContainer.Dock = System.Windows.Forms.DockStyle.Top
        Me.ButtonContainer.Location = New System.Drawing.Point(0, 24)
        Me.ButtonContainer.Name = "ButtonContainer"
        Me.ButtonContainer.Size = New System.Drawing.Size(1043, 71)
        Me.ButtonContainer.TabIndex = 22
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.Panel1.Controls.Add(Me.lblActualizacion)
        Me.Panel1.Location = New System.Drawing.Point(852, 24)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(196, 71)
        Me.Panel1.TabIndex = 216
        '
        'lblActualizacion
        '
        Me.lblActualizacion.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblActualizacion.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.lblActualizacion.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblActualizacion.ForeColor = System.Drawing.Color.Black
        Me.lblActualizacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblActualizacion.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblActualizacion.Location = New System.Drawing.Point(50, 1)
        Me.lblActualizacion.Name = "lblActualizacion"
        Me.lblActualizacion.Size = New System.Drawing.Size(125, 66)
        Me.lblActualizacion.TabIndex = 217
        Me.lblActualizacion.Text = "Cambios en la Version!" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Click para ver los cambios"
        Me.lblActualizacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'panel2
        '
        Me.panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.panel2.Controls.Add(Me.btnmaxi)
        Me.panel2.Controls.Add(Me.lblhora)
        Me.panel2.Controls.Add(Me.lblfecha)
        Me.panel2.Controls.Add(Me.lbluser)
        Me.panel2.Controls.Add(Me.btnmin)
        Me.panel2.Controls.Add(Me.btncerrar)
        Me.panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.panel2.ForeColor = System.Drawing.Color.Gray
        Me.panel2.Location = New System.Drawing.Point(0, 484)
        Me.panel2.Margin = New System.Windows.Forms.Padding(4)
        Me.panel2.Name = "panel2"
        Me.panel2.Size = New System.Drawing.Size(1043, 49)
        Me.panel2.TabIndex = 23
        '
        'btnmaxi
        '
        Me.btnmaxi.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnmaxi.BackColor = System.Drawing.Color.Gray
        Me.btnmaxi.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnmaxi.ForeColor = System.Drawing.Color.White
        Me.btnmaxi.Location = New System.Drawing.Point(901, 2)
        Me.btnmaxi.Margin = New System.Windows.Forms.Padding(4)
        Me.btnmaxi.Name = "btnmaxi"
        Me.btnmaxi.Size = New System.Drawing.Size(48, 44)
        Me.btnmaxi.TabIndex = 3
        Me.btnmaxi.Text = "+"
        Me.btnmaxi.UseVisualStyleBackColor = False
        '
        'lblhora
        '
        Me.lblhora.BackColor = System.Drawing.Color.Transparent
        Me.lblhora.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblhora.ForeColor = System.Drawing.Color.Black
        Me.lblhora.Location = New System.Drawing.Point(4, 12)
        Me.lblhora.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblhora.Name = "lblhora"
        Me.lblhora.Size = New System.Drawing.Size(175, 25)
        Me.lblhora.TabIndex = 0
        Me.lblhora.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblfecha
        '
        Me.lblfecha.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblfecha.BackColor = System.Drawing.Color.Transparent
        Me.lblfecha.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblfecha.ForeColor = System.Drawing.Color.Black
        Me.lblfecha.Location = New System.Drawing.Point(776, 8)
        Me.lblfecha.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblfecha.Name = "lblfecha"
        Me.lblfecha.Size = New System.Drawing.Size(117, 31)
        Me.lblfecha.TabIndex = 2
        Me.lblfecha.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbluser
        '
        Me.lbluser.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbluser.BackColor = System.Drawing.Color.Transparent
        Me.lbluser.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbluser.ForeColor = System.Drawing.Color.Black
        Me.lbluser.Location = New System.Drawing.Point(297, 7)
        Me.lbluser.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbluser.Name = "lbluser"
        Me.lbluser.Size = New System.Drawing.Size(329, 34)
        Me.lbluser.TabIndex = 2
        Me.lbluser.Text = "Usuario: Angel "
        Me.lbluser.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnmin
        '
        Me.btnmin.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnmin.BackColor = System.Drawing.Color.Gray
        Me.btnmin.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnmin.ForeColor = System.Drawing.Color.White
        Me.btnmin.Location = New System.Drawing.Point(948, 2)
        Me.btnmin.Margin = New System.Windows.Forms.Padding(4)
        Me.btnmin.Name = "btnmin"
        Me.btnmin.Size = New System.Drawing.Size(48, 44)
        Me.btnmin.TabIndex = 1
        Me.btnmin.Text = "-"
        Me.btnmin.UseVisualStyleBackColor = False
        '
        'btncerrar
        '
        Me.btncerrar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btncerrar.BackColor = System.Drawing.Color.Gray
        Me.btncerrar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btncerrar.ForeColor = System.Drawing.Color.White
        Me.btncerrar.Location = New System.Drawing.Point(995, 2)
        Me.btncerrar.Margin = New System.Windows.Forms.Padding(4)
        Me.btncerrar.Name = "btncerrar"
        Me.btncerrar.Size = New System.Drawing.Size(48, 44)
        Me.btncerrar.TabIndex = 2
        Me.btncerrar.Text = "x"
        Me.btncerrar.UseVisualStyleBackColor = False
        '
        'MDIMenuPrincipalCorpo
        '
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1043, 533)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.panel2)
        Me.Controls.Add(Me.ButtonContainer)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "MDIMenuPrincipalCorpo"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Corporativo - Joyería y Regalos S.A."
        CType(Me.MenuAcercaDe, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.menuContextualGenOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuArchivoOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuBancosOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuBancosOpcCatalogos, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuBancosOpcProcesoDiario, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuBancosOpcProcesoDiarioRptOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuBancosOpcProcesoMensual, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuCatalogosOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuCompyCxPOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuCompyCxPRptOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuConfiguracionOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuConfiguracionOpcUtil, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuContextualOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuEdicionOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuFacturacionOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuFacturacionRptFactOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuInvEntradasOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuInvHojaOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuInvOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuInvRptOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuInvSalidasOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuSegOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuVentasOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuVentasSalMerOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuVentasSalMerOpcRepEjec, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuVentasVendExtOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuVentasVtasIngrOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuVerOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuVerToolBarOpc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuVerVentanaOpc, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ButtonContainer.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.panel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

End Class