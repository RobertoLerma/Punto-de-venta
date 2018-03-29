Option Explicit On
Option Strict Off
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class FrmConsultas

    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip

    Dim RenAnt As Integer
    Dim I As Integer
    Dim VarAux As Object
    Dim CodArt As Integer
    Public WithEvents Flexdet As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Dim Grid As Integer
    Public bandera As Boolean = False

    Private Sub FlexDet_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Flexdet.DblClick
        Aceptar()
    End Sub

    Private Sub FlexDet_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles Flexdet.KeyPressEvent
        If eventArgs.keyAscii = System.Windows.Forms.Keys.Return Then Aceptar()
    End Sub

    Private Sub FrmConsultas_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Me.Flexdet.Row = 1
        Me.Flexdet.Col = 0
        ModEstandar.CentrarForma(Me)
    End Sub

    Public Sub Aceptar()
        On Error GoTo Merr
        Dim strCtaBancaria As String
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Dim Columna As Integer
        With Flexdet
            Select Case Me.Tag
                'ABC a Bancos (Rosaura)
                Case "FRMCORPOABCBANCOS.TXTCODBANCO"
                    With FrmAbcBancos
                        .txtCodBanco.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMCORPOABCBANCOS.TXTDESCRIPCION"
                    With FrmAbcBancos
                        .txtCodBanco.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                    'ABC A Tipos de MAterial (Rossaura)
                Case "FRMCORPOABCTIPOSMATERIAL.TXTCODTIPOMATERIAL"
                    With frmCorpoAbcTiposMaterial
                        .txtCodTipoMaterial.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMCORPOABCTIPOSMATERIAL.TXTDESCRIPCION"
                    With frmCorpoAbcTiposMaterial
                        .txtCodTipoMaterial.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                    'Fin a ABC  a Tipos de Material
                    'ABC a TALLERES
                Case "FRMCORPOABCTALLERES.TXTCODTALLER"
                    With frmCorpoAbcTalleres
                        .txtCodTaller.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMCORPOABCTALLERES.TXTDESCRIPCION"
                    With frmCorpoAbcTalleres
                        .txtCodTaller.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                    'FIN a ABC  TALLERES
                    'ABC a VENDEDORES
                Case "FRMCORPOABCVENDEDORES.TXTCODVENDEDOR"
                    With FrmCorpoAbcVendedores
                        .txtCodVendedor.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMCORPOABCVENDEDORES.TXTDESCRIPCION"
                    With FrmCorpoAbcVendedores
                        .txtCodVendedor.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                    'FIN a ABC  VENDEDORES
                    'ABC a PROV/ACREED
                Case "FRMCORPOABCPROVACREED.TXTCODPROVACREED"
                    With frmCorpoAbcProvAcreed
                        .txtCodProvAcreed.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMCORPOABCPROVACREED.TXTNOMBRE"
                    With frmCorpoAbcProvAcreed
                        .txtCodProvAcreed.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                    'FIN a ABC  PROV/ACREEED
                    'ABC a sucursales
                Case "FRMCORPOABCSUCURSALES.TXTCODSUCURSAL"
                    With frmCorpoAbcSucursales
                        .txtCodSucursal.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMCORPOABCSUCURSALES.TXTDESCRIPCION"
                    With frmCorpoAbcSucursales
                        .txtCodSucursal.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                    'FIN a ABC  sucursales
                    'ABC a Formas de Pago
                Case "FRMCORPOABCFORMASDEPAGO.TXTCODFORMAPAGO"
                    With frmCorpoAbcFormasdePago
                        .txtCodFormaPago.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMCORPOABCFORMASDEPAGO.TXTDESCRIPCION"
                    With frmCorpoAbcFormasdePago
                        .txtCodFormaPago.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                    'Fin ABC a Formas de Pago
                Case "FRMABCFUNCIONES.TXTCODFUNCION"
                    With frmABCFunciones
                        .txtCodFuncion.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMABCFUNCIONES.TXTDESCFUNCION"
                    With frmABCFunciones
                        .txtCodFuncion.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMABCMODULOS.TXTCODMODULO"
                    With frmABCModulos
                        .txtCodModulo.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMABCMODULOS.TXTDESCMODULO"
                    With frmABCModulos
                        .txtCodModulo.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                    '--------------------------------------------------------------------------------------------------------------
                    ' CÓDIGO DE PAIMÍ
                    '--------------------------------------------------------------------------------------------------------------
                Case "FRMCORPOABCARTICULOS.TXTCODARTICULO", "FRMCORPOABCARTICULOS.TXTDESCARTICULO", "FRMCORPOABCARTICULOS.TXTCODIGODELPROVEEDOR"
                    With frmCorpoABCArticulos
                        .txtCodArticulo.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                        Exit Sub
                    End With
                Case "FRMABCUSUARIOS.TXTCODIGO"
                    With frmABCUsuarios
                        .txtCodigo.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMABCUSUARIOS.TXTNOMBRE"
                    With frmABCUsuarios
                        .txtCodigo.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                    '            Case "FRMCORPOABCARTICULOS.TXTDESCRIPCION"
                    '                With frmCorpoABCArticulos
                    '                    .txtCodArticulo.text = Flexdet.TextMatrix(Me.Flexdet.Row, 1)
                    '                    .LlenaDatos
                    '                    Unload Me
                    '                End With
                Case "FRMCORPOABCGRUPOS.TXTCODGRUPO"
                    With frmCorpoABCGrupos
                        .txtCodGrupo.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMCORPOABCGRUPOS.TXTDESCGRUPO"
                    With frmCorpoABCGrupos
                        .txtCodGrupo.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMCORPOABCMODELOS.TXTCODGRUPO"
                    With frmCorpoABCModelos
                        '.txtCodGrupo.text = Flexdet.TextMatrix(Me.Flexdet.Row, 0)
                        '.LlenaDatos
                        'Unload Me
                    End With
                Case "FRMCORPOABCMODELOS.TXTDESCGRUPO"
                    With frmCorpoABCModelos
                        '.txtCodGrupo.text = Flexdet.TextMatrix(Me.Flexdet.Row, 1)
                        '.LlenaDatos
                        'Unload Me
                    End With
                Case "FRMCORPOABCCUENTASBANCARIAS.TXTCTABANCARIA"
                    With frmAbcCuentasBancarias
                        strCtaBancaria = Trim(Flexdet.get_TextMatrix(Me.Flexdet.Row, 0))
                        .txtCtaBancaria.Text = strCtaBancaria
                        .mintCodBanco = CInt(ModEstandar.Numerico(Flexdet.get_TextMatrix(Me.Flexdet.Row, 3)))
                        .LlenaDatos()
                        Me.Close()
                        .txtCtaBancaria.Focus()
                    End With
                Case "FRMCORPOABCCUENTASBANCARIAS.TXTCUENTAHABIENTE"
                    With frmAbcCuentasBancarias
                        strCtaBancaria = Trim(Flexdet.get_TextMatrix(Me.Flexdet.Row, 0))
                        .txtCtaBancaria.Text = strCtaBancaria
                        .mintCodBanco = CInt(ModEstandar.Numerico(Flexdet.get_TextMatrix(Me.Flexdet.Row, 3)))
                        .LlenaDatos()
                        Me.Close()
                        .txtCtaBancaria.Focus()
                    End With
                Case "FRMCORPOABCCUENTASBANCARIAS.TXTSUCURSAL"
                    With frmAbcCuentasBancarias
                        .txtSucursal.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 1)
                        .mintCodBanco = CInt(ModEstandar.Numerico(Flexdet.get_TextMatrix(Me.Flexdet.Row, 3)))
                        .LlenaDatos()
                        Me.Close()
                        .txtCtaBancaria.Focus()
                    End With
                Case "FRMCXPORDENCOMPRA.MSHFLEX"
                    '            With frmCXPOrdenCompra
                    '                If .mshFlex.Col = 0 Then
                    '                    .mshFlex.TextMatrix(.mshFlex.Row, 0) = Flexdet.TextMatrix(Me.Flexdet.Row, 0)
                    '                ElseIf .mshFlex.Col = 1 Then
                    '                    .mshFlex.TextMatrix(.mshFlex.Row, 0) = Flexdet.TextMatrix(Me.Flexdet.Row, 0)
                    '                End If
                    '                .LlenaLineaGrid
                    '                Unload Me
                    '                .mshFlex.SetFocus
                    '                .ActualizaCantidades
                    '                Exit Sub
                    '            End With
                Case "FRMCXPORDENCOMPRA.TXTFLEX"
                    With frmCXPOrdenCompra
                        'If .mshFlex.Col = 0 Then
                        '    .txtFlex.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        'ElseIf .mshFlex.Col = 1 Then
                        '    .txtFlex.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        'End If
                        '.ArticuloRepetido(CInt(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        'Me.Close()
                        '.mshFlex.Focus()
                        '.ActualizaCantidades()
                        'Exit Sub
                    End With
                Case "FRMCXPREGFACTCOMPRAS.TXTFOLIOFACTURA"
                    'With frmCXPRegFactCompras
                    '    .txtFolioFactura.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                    '    .mintCodProveedor = CInt(Numerico(Flexdet.get_TextMatrix(Me.Flexdet.Row, 5)))
                    '    .nNUMDOCTO = CInt(Numerico(Flexdet.get_TextMatrix(Me.Flexdet.Row, 6)))
                    '    .lOperExt = True
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMCXPREGFACTCOMPRASCARGAINICIAL.TXTFOLIOFACTURA"
                    'With frmCXPRegFactComprasCargaInicial
                    '    .txtFolioFactura.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                    '    .mintCodProveedor = CInt(Numerico(Flexdet.get_TextMatrix(Me.Flexdet.Row, 5)))
                    '    .nNUMDOCTO = CInt(Numerico(Flexdet.get_TextMatrix(Me.Flexdet.Row, 6)))
                    '    .lOperExt = True
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMCXPREGFACTGASTOS.TXTFOLIOFACTURA"
                    'With frmCXPRegFactGastos
                    '    .txtFolioFactura.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                    '    .mintCodProveedor = CInt(Numerico(Flexdet.get_TextMatrix(Me.Flexdet.Row, 5)))
                    '    .nNUMDOCTO = CInt(Numerico(Flexdet.get_TextMatrix(Me.Flexdet.Row, 6)))
                    '    .lOperExt = True
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMCXPREGNOTASCREDITO.TXTFOLIO"
                    'With frmCXPRegNotasCredito
                    '    .txtFolio(.sstNota.SelectedIndex).Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMCXPREGNOTASCREDITO.TXTFACTURA"
                    'With frmCXPRegNotasCredito
                    '    .txtFactura(.sstNota.SelectedIndex).Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                    '    .nNUMDOCTO = CInt(Numerico(Flexdet.get_TextMatrix(Me.Flexdet.Row, 6)))
                    '    .LlenaFactura()
                    '    Me.Close()
                    'End With
                    '--------------------------------------------------------------------------------------------------------------
                    ' FINALIZA CÓDIGO DE PAIMÍ
                    '--------------------------------------------------------------------------------------------------------------
                    'ABC de Origen y Aplicación de Recursos
                Case "FRMCORPOABCORIGENYAPLICACIONDERECURSOS.TXTCODIGO"
                    With frmAbcOrigenyAplicaciondeRecursos
                        .txtCodigo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMCORPOABCORIGENYAPLICACIONDERECURSOS.TXTDESCRIPCION"
                    With frmAbcOrigenyAplicaciondeRecursos
                        .txtCodigo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                    'Fin de ABC de Origen y Aplicación de Recursos
                    'ABC de Rubros de Origen y Aplicación
                Case "FRMCORPOABCRUBROSDEAPLICACIONYORIGEN.TXTCODIGO"
                    With frmCorpoABCRubrosdeAplicacionyOrigen
                        .txtCodigo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMCORPOABCRUBROSDEAPLICACIONYORIGEN.TXTDESCRIPCION"
                    With frmCorpoABCRubrosdeAplicacionyOrigen
                        .txtCodigo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                    'Fin ABC de Rubros de Origen y Aplicación
                    'ABC de Clientes
                Case "FRMCORPOABCCLIENTES.TXTCODIGO"
                    With frmCorpoABCClientes
                        .txtCodigo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMCORPOABCCLIENTES.TXTNOMBRE"
                    With frmCorpoABCClientes
                        .txtCodigo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                    'Fin de ABC de Clientes
                    'Busqueda de Clientes en Facturación Especial
                Case "FRMFACTFACTURACIONESPECIAL.TXTCODIGO"
                    With frmFactFacturacionEspecial
                        .txtCodigo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        .LlenaDatosCliente()
                        Me.Close()
                    End With
                Case "FRMFACTFACTURACIONESPECIAL.TXTNOMBRECLIENTE"
                    With frmFactFacturacionEspecial
                        .txtCodigo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                        .LlenaDatosCliente()
                        Me.Close()
                    End With
                Case "FRMFACTFACTURACIONESPECIAL.TXTFOLIOFACTURA"
                    With frmFactFacturacionEspecial
                        .txtFolioFactura.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        .LlenaDatos()
                        Me.Close()
                    End With
                    'Fin de Busqueda en Facturación Especial
                Case "FRMPAGOS.CODIGO AGRUPADOR"
                    'With frmPagos.flexDetalle
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    If Trim(frmPagos.flexDetalle.get_TextMatrix(frmPagos.flexDetalle.Row, 0)) <> Trim(frmPagos.txtFlex.Text) Then
                    '        frmPagos.flexDetalle.set_TextMatrix(frmPagos.flexDetalle.Row, 2, "")
                    '        frmPagos.flexDetalle.set_TextMatrix(frmPagos.flexDetalle.Row, 3, "")
                    '        frmPagos.flexDetalle.set_TextMatrix(frmPagos.flexDetalle.Row, 4, "")
                    '    End If
                    '    frmPagos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmPagos.ValidaLlave()
                    'End With
                Case "FRMPAGOS.CODIGO RUBRO"
                    'With frmPagos.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    frmPagos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmPagos.ValidaLlave()
                    'End With
                Case "FRMPAGOS.DESCRIPCION AGRUPADOR"
                    'With frmPagos.flexDetalle
                    '    If Trim(frmPagos.flexDetalle.get_TextMatrix(frmPagos.flexDetalle.Row, 1)) <> Trim(frmPagos.txtFlex.Text) Then
                    '        frmPagos.flexDetalle.set_TextMatrix(frmPagos.flexDetalle.Row, 2, "")
                    '        frmPagos.flexDetalle.set_TextMatrix(frmPagos.flexDetalle.Row, 3, "")
                    '        frmPagos.flexDetalle.set_TextMatrix(frmPagos.flexDetalle.Row, 4, "")
                    '    End If
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmPagos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmPagos.ValidaLlave()
                    'End With
                Case "FRMPAGOS.DESCRIPCION RUBRO"
                    'With frmPagos.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmPagos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmPagos.ValidaLlave()
                    'End With

                Case "FRMDEPOSITOS.CODIGO AGRUPADOR"
                    'With frmDepositos.flexDetalle
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    If Trim(frmDepositos.flexDetalle.get_TextMatrix(frmDepositos.flexDetalle.Row, 0)) <> Trim(frmDepositos.txtFlex.Text) Then
                    '        frmDepositos.flexDetalle.set_TextMatrix(frmDepositos.flexDetalle.Row, 2, "")
                    '        frmDepositos.flexDetalle.set_TextMatrix(frmDepositos.flexDetalle.Row, 3, "")
                    '        frmDepositos.flexDetalle.set_TextMatrix(frmDepositos.flexDetalle.Row, 4, "")
                    '    End If
                    '    frmDepositos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmDepositos.ValidaLlave()
                    'End With
                Case "FRMDEPOSITOSINTPES.CODIGO AGRUPADOR"
                    'With frmDepositosIntPes.flexDetalle
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    If Trim(frmDepositosIntPes.flexDetalle.get_TextMatrix(frmDepositosIntPes.flexDetalle.Row, 0)) <> Trim(frmDepositosIntPes.txtFlex.Text) Then
                    '        frmDepositosIntPes.flexDetalle.set_TextMatrix(frmDepositosIntPes.flexDetalle.Row, 2, "")
                    '        frmDepositosIntPes.flexDetalle.set_TextMatrix(frmDepositosIntPes.flexDetalle.Row, 3, "")
                    '        frmDepositosIntPes.flexDetalle.set_TextMatrix(frmDepositosIntPes.flexDetalle.Row, 4, "")
                    '    End If
                    '    frmDepositosIntPes.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmDepositosIntPes.ValidaLlave()
                    'End With
                Case "FRMDEPOSITOSINTDOL.CODIGO AGRUPADOR"
                    'With frmDepositosIntDol.flexDetalle
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    If Trim(frmDepositosIntDol.flexDetalle.get_TextMatrix(frmDepositosIntDol.flexDetalle.Row, 0)) <> Trim(frmDepositosIntDol.txtFlex.Text) Then
                    '        frmDepositosIntDol.flexDetalle.set_TextMatrix(frmDepositosIntDol.flexDetalle.Row, 2, "")
                    '        frmDepositosIntDol.flexDetalle.set_TextMatrix(frmDepositosIntDol.flexDetalle.Row, 3, "")
                    '        frmDepositosIntDol.flexDetalle.set_TextMatrix(frmDepositosIntDol.flexDetalle.Row, 4, "")
                    '    End If
                    '    frmDepositosIntDol.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmDepositosIntDol.ValidaLlave()
                    'End With

                Case "FRMDEPOSITOS.CODIGO RUBRO"
                    'With frmDepositos.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    frmDepositos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmDepositos.ValidaLlave()
                    'End With
                Case "FRMDEPOSITOSINTPES.CODIGO RUBRO"
                    'With frmDepositosIntPes.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    frmDepositosIntPes.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmDepositosIntPes.ValidaLlave()
                    'End With
                Case "FRMDEPOSITOSINTDOL.CODIGO RUBRO"
                    'With frmDepositosIntDol.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    frmDepositosIntDol.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmDepositosIntDol.ValidaLlave()
                    'End With

                Case "FRMDEPOSITOS.DESCRIPCION AGRUPADOR"
                    'With frmDepositos.flexDetalle
                    '    If Trim(frmDepositos.flexDetalle.get_TextMatrix(frmDepositos.flexDetalle.Row, 1)) <> Trim(frmDepositos.txtFlex.Text) Then
                    '        frmDepositos.flexDetalle.set_TextMatrix(frmDepositos.flexDetalle.Row, 2, "")
                    '        frmDepositos.flexDetalle.set_TextMatrix(frmDepositos.flexDetalle.Row, 3, "")
                    '        frmDepositos.flexDetalle.set_TextMatrix(frmDepositos.flexDetalle.Row, 4, "")
                    '    End If
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmDepositos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmDepositos.ValidaLlave()
                    'End With
                Case "FRMDEPOSITOSINTPES.DESCRIPCION AGRUPADOR"
                    'With frmDepositosIntPes.flexDetalle
                    '    If Trim(frmDepositosIntPes.flexDetalle.get_TextMatrix(frmDepositosIntPes.flexDetalle.Row, 1)) <> Trim(frmDepositosIntPes.txtFlex.Text) Then
                    '        frmDepositosIntPes.flexDetalle.set_TextMatrix(frmDepositosIntPes.flexDetalle.Row, 2, "")
                    '        frmDepositosIntPes.flexDetalle.set_TextMatrix(frmDepositosIntPes.flexDetalle.Row, 3, "")
                    '        frmDepositosIntPes.flexDetalle.set_TextMatrix(frmDepositosIntPes.flexDetalle.Row, 4, "")
                    '    End If
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmDepositosIntPes.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmDepositosIntPes.ValidaLlave()
                    'End With
                Case "FRMDEPOSITOSINTDOL.DESCRIPCION AGRUPADOR"
                    'With frmDepositosIntDol.flexDetalle
                    '    If Trim(frmDepositosIntDol.flexDetalle.get_TextMatrix(frmDepositosIntDol.flexDetalle.Row, 1)) <> Trim(frmDepositosIntDol.txtFlex.Text) Then
                    '        frmDepositosIntDol.flexDetalle.set_TextMatrix(frmDepositosIntDol.flexDetalle.Row, 2, "")
                    '        frmDepositosIntDol.flexDetalle.set_TextMatrix(frmDepositosIntDol.flexDetalle.Row, 3, "")
                    '        frmDepositosIntDol.flexDetalle.set_TextMatrix(frmDepositosIntDol.flexDetalle.Row, 4, "")
                    '    End If
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmDepositosIntDol.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmDepositosIntDol.ValidaLlave()
                    'End With

                Case "FRMDEPOSITOS.DESCRIPCION RUBRO"
                    'With frmDepositos.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmDepositos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmDepositos.ValidaLlave()
                    'End With
                Case "FRMDEPOSITOSINTPES.DESCRIPCION RUBRO"
                    'With frmDepositosIntPes.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmDepositosIntPes.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmDepositosIntPes.ValidaLlave()
                    'End With
                Case "FRMDEPOSITOSINTDOL.DESCRIPCION RUBRO"
                    'With frmDepositosIntDol.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmDepositosIntDol.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmDepositosIntDol.ValidaLlave()
                    'End With

                Case "FRMCARGOS.CODIGO AGRUPADOR"
                    'With frmCargos.flexDetalle
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    If Trim(frmCargos.flexDetalle.get_TextMatrix(frmCargos.flexDetalle.Row, 0)) <> Trim(frmCargos.txtFlex.Text) Then
                    '        frmCargos.flexDetalle.set_TextMatrix(frmCargos.flexDetalle.Row, 2, "")
                    '        frmCargos.flexDetalle.set_TextMatrix(frmCargos.flexDetalle.Row, 3, "")
                    '        frmCargos.flexDetalle.set_TextMatrix(frmCargos.flexDetalle.Row, 4, "0.00")
                    '    End If
                    '    frmCargos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmCargos.ValidaLlave()
                    'End With
                Case "FRMCARGOS.CODIGO RUBRO"
                    'With frmCargos.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    frmCargos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmCargos.ValidaLlave()
                    'End With
                Case "FRMCARGOS.DESCRIPCION AGRUPADOR"
                    'With frmCargos.flexDetalle
                    '    If Trim(frmCargos.flexDetalle.get_TextMatrix(frmCargos.flexDetalle.Row, 1)) <> Trim(frmCargos.txtFlex.Text) Then
                    '        frmCargos.flexDetalle.set_TextMatrix(frmCargos.flexDetalle.Row, 2, "")
                    '        frmCargos.flexDetalle.set_TextMatrix(frmCargos.flexDetalle.Row, 3, "")
                    '        frmCargos.flexDetalle.set_TextMatrix(frmCargos.flexDetalle.Row, 4, "0.00")
                    '    End If
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmCargos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmCargos.ValidaLlave()
                    'End With
                Case "FRMCARGOS.DESCRIPCION RUBRO"
                    'With frmCargos.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmCargos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmCargos.ValidaLlave()
                    'End With
                Case "FRMANTICIPOS.CODIGO AGRUPADOR"
                    'With frmAnticipos.flexDetalle
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    If Trim(frmAnticipos.flexDetalle.get_TextMatrix(frmAnticipos.flexDetalle.Row, 0)) <> Trim(frmAnticipos.txtFlex.Text) Then
                    '        frmAnticipos.flexDetalle.set_TextMatrix(frmAnticipos.flexDetalle.Row, 2, "")
                    '        frmAnticipos.flexDetalle.set_TextMatrix(frmAnticipos.flexDetalle.Row, 3, "")
                    '        frmAnticipos.flexDetalle.set_TextMatrix(frmAnticipos.flexDetalle.Row, 4, "0.00")
                    '    End If
                    '    frmAnticipos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmAnticipos.ValidaLlave()
                    'End With
                Case "FRMANTICIPOS.CODIGO RUBRO"
                    'With frmAnticipos.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    frmAnticipos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmAnticipos.ValidaLlave()
                    'End With
                Case "FRMANTICIPOS.DESCRIPCION AGRUPADOR"
                    'With frmAnticipos.flexDetalle
                    '    If Trim(frmAnticipos.flexDetalle.get_TextMatrix(frmAnticipos.flexDetalle.Row, 1)) <> Trim(frmAnticipos.txtFlex.Text) Then
                    '        frmAnticipos.flexDetalle.set_TextMatrix(frmAnticipos.flexDetalle.Row, 2, "")
                    '        frmAnticipos.flexDetalle.set_TextMatrix(frmAnticipos.flexDetalle.Row, 3, "")
                    '        frmAnticipos.flexDetalle.set_TextMatrix(frmAnticipos.flexDetalle.Row, 4, "0.00")
                    '    End If
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmAnticipos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmAnticipos.ValidaLlave()
                    'End With
                Case "FRMANTICIPOS.DESCRIPCION RUBRO"
                    'With frmAnticipos.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmAnticipos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmAnticipos.ValidaLlave()
                    'End With
                Case "FRMOTROSINGRESOS.CODIGO AGRUPADOR"
                    'With frmOtrosIngresos.flexDetalle
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    If Trim(frmOtrosIngresos.flexDetalle.get_TextMatrix(frmOtrosIngresos.flexDetalle.Row, 0)) <> Trim(frmOtrosIngresos.txtFlex.Text) Then
                    '        frmOtrosIngresos.flexDetalle.set_TextMatrix(frmOtrosIngresos.flexDetalle.Row, 2, "")
                    '        frmOtrosIngresos.flexDetalle.set_TextMatrix(frmOtrosIngresos.flexDetalle.Row, 3, "")
                    '        frmOtrosIngresos.flexDetalle.set_TextMatrix(frmOtrosIngresos.flexDetalle.Row, 4, "0.00")
                    '    End If
                    '    frmOtrosIngresos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmOtrosIngresos.ValidaLlave()
                    'End With
                Case "FRMOTROSINGRESOS.CODIGO RUBRO"
                    'With frmOtrosIngresos.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    frmOtrosIngresos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmOtrosIngresos.ValidaLlave()
                    'End With
                Case "FRMOTROSINGRESOS.DESCRIPCION AGRUPADOR"
                    'With frmOtrosIngresos.flexDetalle
                    '    If Trim(frmOtrosIngresos.flexDetalle.get_TextMatrix(frmOtrosIngresos.flexDetalle.Row, 1)) <> Trim(frmOtrosIngresos.txtFlex.Text) Then
                    '        frmOtrosIngresos.flexDetalle.set_TextMatrix(frmOtrosIngresos.flexDetalle.Row, 2, "")
                    '        frmOtrosIngresos.flexDetalle.set_TextMatrix(frmOtrosIngresos.flexDetalle.Row, 3, "")
                    '        frmOtrosIngresos.flexDetalle.set_TextMatrix(frmOtrosIngresos.flexDetalle.Row, 4, "0.00")
                    '    End If
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmOtrosIngresos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmOtrosIngresos.ValidaLlave()
                    'End With
                Case "FRMOTROSINGRESOS.DESCRIPCION RUBRO"
                    'With frmOtrosIngresos.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmOtrosIngresos.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmOtrosIngresos.ValidaLlave()
                    'End With
                Case "FRMCONSULTAORIGENAPLICACION.CODIGO AGRUPADOR"
                    'With frmConsultaOrigenAplicacion.flexDetalle
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    If Trim(frmConsultaOrigenAplicacion.flexDetalle.get_TextMatrix(frmConsultaOrigenAplicacion.flexDetalle.Row, 0)) <> Trim(frmConsultaOrigenAplicacion.txtFlex.Text) Then
                    '        frmConsultaOrigenAplicacion.flexDetalle.set_TextMatrix(frmConsultaOrigenAplicacion.flexDetalle.Row, 2, "")
                    '        frmConsultaOrigenAplicacion.flexDetalle.set_TextMatrix(frmConsultaOrigenAplicacion.flexDetalle.Row, 3, "")
                    '        frmConsultaOrigenAplicacion.flexDetalle.set_TextMatrix(frmConsultaOrigenAplicacion.flexDetalle.Row, 4, "0.00")
                    '    End If
                    '    frmConsultaOrigenAplicacion.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmConsultaOrigenAplicacion.ValidaLlave()
                    'End With
                Case "FRMCONSULTAORIGENAPLICACION.CODIGO RUBRO"
                    'With frmConsultaOrigenAplicacion.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    frmConsultaOrigenAplicacion.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    frmConsultaOrigenAplicacion.ValidaLlave()
                    'End With
                Case "FRMCONSULTAORIGENAPLICACION.DESCRIPCION AGRUPADOR"
                    'With frmConsultaOrigenAplicacion.flexDetalle
                    '    If Trim(frmConsultaOrigenAplicacion.flexDetalle.get_TextMatrix(frmConsultaOrigenAplicacion.flexDetalle.Row, 1)) <> Trim(frmConsultaOrigenAplicacion.txtFlex.Text) Then
                    '        frmConsultaOrigenAplicacion.flexDetalle.set_TextMatrix(frmConsultaOrigenAplicacion.flexDetalle.Row, 2, "")
                    '        frmConsultaOrigenAplicacion.flexDetalle.set_TextMatrix(frmConsultaOrigenAplicacion.flexDetalle.Row, 3, "")
                    '        frmConsultaOrigenAplicacion.flexDetalle.set_TextMatrix(frmConsultaOrigenAplicacion.flexDetalle.Row, 4, "0.00")
                    '    End If
                    '    .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmConsultaOrigenAplicacion.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmConsultaOrigenAplicacion.ValidaLlave()
                    'End With
                Case "FRMCONSULTAORIGENAPLICACION.DESCRIPCION RUBRO"
                    'With frmConsultaOrigenAplicacion.flexDetalle
                    '    .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    frmConsultaOrigenAplicacion.txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    Me.Close()
                    '    frmConsultaOrigenAplicacion.ValidaLlave()
                    'End With
                    'Principia Busquedas para el estado de resultados
                Case "EDORESULTADOS.DESCRIPCION SUCURSAL"
                    With frmVtasEstadodeResultados.flexGastos
                        .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        .set_TextMatrix(.Row, 5, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                        Me.Close()
                        If frmVtasEstadodeResultados.ValidaCodigos((frmVtasEstadodeResultados.flexGastos.Row)) Then
                            MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            frmVtasEstadodeResultados.flexGastos.Col = 0
                            frmVtasEstadodeResultados.flexGastos.Focus()
                        End If
                        frmVtasEstadodeResultados.flexGastos.Col = 1
                        frmVtasEstadodeResultados.flexGastos.Focus()
                    End With
                Case "EDORESULTADOS.CODIGO AGRUPADOR"
                    With frmVtasEstadodeResultados.flexGastos
                        .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        'frmPagos.txtFlex = Trim(Flexdet.TextMatrix(Flexdet.Row, 0))
                        .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                        Me.Close()
                        If frmVtasEstadodeResultados.ValidaCodigos((frmVtasEstadodeResultados.flexGastos.Row)) Then
                            MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            frmVtasEstadodeResultados.flexGastos.set_TextMatrix(frmVtasEstadodeResultados.flexGastos.Row, 1, "")
                            frmVtasEstadodeResultados.flexGastos.set_TextMatrix(frmVtasEstadodeResultados.flexGastos.Row, 2, "")
                            frmVtasEstadodeResultados.flexGastos.Col = 1
                            frmVtasEstadodeResultados.flexGastos.Focus()
                        Else
                            frmVtasEstadodeResultados.flexGastos.set_TextMatrix(frmVtasEstadodeResultados.flexGastos.Row, 3, "")
                            frmVtasEstadodeResultados.flexGastos.set_TextMatrix(frmVtasEstadodeResultados.flexGastos.Row, 4, "")
                            If frmVtasEstadodeResultados.flexGastos.Row = frmVtasEstadodeResultados.flexGastos.Rows - 1 Then
                                frmVtasEstadodeResultados.flexGastos.Rows = frmVtasEstadodeResultados.flexGastos.Rows + 1
                            End If
                        End If
                        frmVtasEstadodeResultados.flexGastos.Col = 3
                        frmVtasEstadodeResultados.flexGastos.Focus()
                    End With
                Case "EDORESULTADOS.CODIGO RUBRO"
                    With frmVtasEstadodeResultados.flexGastos
                        .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        .set_TextMatrix(.Row, 4, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                        Me.Close()
                        If frmVtasEstadodeResultados.ValidaCodigos((frmVtasEstadodeResultados.flexGastos.Row)) Then
                            MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            frmVtasEstadodeResultados.flexGastos.set_TextMatrix(frmVtasEstadodeResultados.flexGastos.Row, 3, "")
                            frmVtasEstadodeResultados.flexGastos.set_TextMatrix(frmVtasEstadodeResultados.flexGastos.Row, 4, "")
                            frmVtasEstadodeResultados.flexGastos.Col = 3
                            frmVtasEstadodeResultados.flexGastos.Focus()
                        Else
                            If frmVtasEstadodeResultados.flexGastos.Row = frmVtasEstadodeResultados.flexGastos.Rows - 1 Then
                                frmVtasEstadodeResultados.flexGastos.Rows = frmVtasEstadodeResultados.flexGastos.Rows + 1
                                frmVtasEstadodeResultados.flexGastos.Row = frmVtasEstadodeResultados.flexGastos.Row + 1
                                frmVtasEstadodeResultados.flexGastos.TopRow = frmVtasEstadodeResultados.flexGastos.Row
                            Else
                                frmVtasEstadodeResultados.flexGastos.Row = frmVtasEstadodeResultados.flexGastos.Row + 1
                                If frmVtasEstadodeResultados.flexGastos.Row > 6 Then
                                    frmVtasEstadodeResultados.flexGastos.TopRow = frmVtasEstadodeResultados.flexGastos.Row
                                End If
                            End If
                        End If
                        frmVtasEstadodeResultados.flexGastos.Col = 0
                        frmVtasEstadodeResultados.flexGastos.Focus()
                    End With
                Case "EDORESULTADOS.DESCRIPCION AGRUPADOR"
                    With frmVtasEstadodeResultados.flexGastos
                        .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                        'frmPagos.txtFlex = Trim(Flexdet.TextMatrix(Flexdet.Row, 0))
                        .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        Me.Close()
                        If frmVtasEstadodeResultados.ValidaCodigos((frmVtasEstadodeResultados.flexGastos.Row)) Then
                            MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            frmVtasEstadodeResultados.flexGastos.set_TextMatrix(frmVtasEstadodeResultados.flexGastos.Row, 1, "")
                            frmVtasEstadodeResultados.flexGastos.set_TextMatrix(frmVtasEstadodeResultados.flexGastos.Row, 2, "")
                            frmVtasEstadodeResultados.flexGastos.Col = 2
                            frmVtasEstadodeResultados.flexGastos.Focus()
                        Else
                            frmVtasEstadodeResultados.flexGastos.set_TextMatrix(frmVtasEstadodeResultados.flexGastos.Row, 3, "")
                            frmVtasEstadodeResultados.flexGastos.set_TextMatrix(frmVtasEstadodeResultados.flexGastos.Row, 4, "")
                            If frmVtasEstadodeResultados.flexGastos.Row = frmVtasEstadodeResultados.flexGastos.Rows - 1 Then
                                frmVtasEstadodeResultados.flexGastos.Rows = frmVtasEstadodeResultados.flexGastos.Rows + 1
                            End If
                        End If
                        frmVtasEstadodeResultados.flexGastos.Col = 3
                        frmVtasEstadodeResultados.flexGastos.Focus()
                    End With
                Case "EDORESULTADOS.DESCRIPCION RUBRO"
                    With frmVtasEstadodeResultados.flexGastos
                        .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                        .set_TextMatrix(.Row, 4, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        Me.Close()
                        If frmVtasEstadodeResultados.ValidaCodigos((frmVtasEstadodeResultados.flexGastos.Row)) Then
                            MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            frmVtasEstadodeResultados.flexGastos.set_TextMatrix(frmVtasEstadodeResultados.flexGastos.Row, 3, "")
                            frmVtasEstadodeResultados.flexGastos.set_TextMatrix(frmVtasEstadodeResultados.flexGastos.Row, 4, "")
                            frmVtasEstadodeResultados.flexGastos.Col = 4
                            frmVtasEstadodeResultados.flexGastos.Focus()
                        Else
                            If frmVtasEstadodeResultados.flexGastos.Row = frmVtasEstadodeResultados.flexGastos.Rows - 1 Then
                                frmVtasEstadodeResultados.flexGastos.Rows = frmVtasEstadodeResultados.flexGastos.Rows + 1
                                frmVtasEstadodeResultados.flexGastos.Row = frmVtasEstadodeResultados.flexGastos.Row + 1
                                frmVtasEstadodeResultados.flexGastos.TopRow = frmVtasEstadodeResultados.flexGastos.Row
                            Else
                                frmVtasEstadodeResultados.flexGastos.Row = frmVtasEstadodeResultados.flexGastos.Row + 1
                                If frmVtasEstadodeResultados.flexGastos.Row > 6 Then
                                    frmVtasEstadodeResultados.flexGastos.TopRow = frmVtasEstadodeResultados.flexGastos.Row
                                End If
                            End If
                        End If
                        frmVtasEstadodeResultados.flexGastos.Col = 0
                        frmVtasEstadodeResultados.flexGastos.Focus()
                    End With
                    'Termina Busquedas del estado de resultados

                    'Principia Busquedas para la relacion de gastos
                Case "RELGASTOS.DESCRIPCION SUCURSAL"
                    With frmVtasRelacionGastos.flexGastos
                        .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        .set_TextMatrix(.Row, 5, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                        Me.Close()
                        If frmVtasRelacionGastos.ValidaCodigos((frmVtasRelacionGastos.flexGastos.Row)) Then
                            MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            frmVtasRelacionGastos.flexGastos.Col = 0
                            frmVtasRelacionGastos.flexGastos.Focus()
                        End If
                        frmVtasRelacionGastos.flexGastos.Col = 1
                        frmVtasRelacionGastos.flexGastos.Focus()
                    End With
                Case "RELGASTOS.CODIGO AGRUPADOR"
                    With frmVtasRelacionGastos.flexGastos
                        .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        'frmPagos.txtFlex = Trim(Flexdet.TextMatrix(Flexdet.Row, 0))
                        .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                        Me.Close()
                        If frmVtasRelacionGastos.ValidaCodigos((frmVtasRelacionGastos.flexGastos.Row)) Then
                            MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            frmVtasRelacionGastos.flexGastos.set_TextMatrix(frmVtasRelacionGastos.flexGastos.Row, 1, "")
                            frmVtasRelacionGastos.flexGastos.set_TextMatrix(frmVtasRelacionGastos.flexGastos.Row, 2, "")
                            frmVtasRelacionGastos.flexGastos.Col = 1
                            frmVtasRelacionGastos.flexGastos.Focus()
                        Else
                            frmVtasRelacionGastos.flexGastos.set_TextMatrix(frmVtasRelacionGastos.flexGastos.Row, 3, "")
                            frmVtasRelacionGastos.flexGastos.set_TextMatrix(frmVtasRelacionGastos.flexGastos.Row, 4, "")
                            If frmVtasRelacionGastos.flexGastos.Row = frmVtasRelacionGastos.flexGastos.Rows - 1 Then
                                frmVtasRelacionGastos.flexGastos.Rows = frmVtasRelacionGastos.flexGastos.Rows + 1
                            End If
                        End If
                        frmVtasRelacionGastos.flexGastos.Col = 3
                        frmVtasRelacionGastos.flexGastos.Focus()
                    End With
                Case "RELGASTOS.CODIGO RUBRO"
                    With frmVtasRelacionGastos.flexGastos
                        .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        .set_TextMatrix(.Row, 4, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                        Me.Close()
                        If frmVtasRelacionGastos.ValidaCodigos((frmVtasRelacionGastos.flexGastos.Row)) Then
                            MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            frmVtasRelacionGastos.flexGastos.set_TextMatrix(frmVtasRelacionGastos.flexGastos.Row, 3, "")
                            frmVtasRelacionGastos.flexGastos.set_TextMatrix(frmVtasRelacionGastos.flexGastos.Row, 4, "")
                            frmVtasRelacionGastos.flexGastos.Col = 3
                            frmVtasRelacionGastos.Activate()
                        Else
                            If frmVtasRelacionGastos.flexGastos.Row = frmVtasRelacionGastos.flexGastos.Rows - 1 Then
                                frmVtasRelacionGastos.flexGastos.Rows = frmVtasRelacionGastos.flexGastos.Rows + 1
                                frmVtasRelacionGastos.flexGastos.Row = frmVtasRelacionGastos.flexGastos.Row + 1
                                frmVtasRelacionGastos.flexGastos.TopRow = frmVtasRelacionGastos.flexGastos.Row
                            Else
                                frmVtasRelacionGastos.flexGastos.Row = frmVtasRelacionGastos.flexGastos.Row + 1
                                If frmVtasRelacionGastos.flexGastos.Row > 6 Then
                                    frmVtasRelacionGastos.flexGastos.TopRow = frmVtasRelacionGastos.flexGastos.Row
                                End If
                            End If
                        End If
                        frmVtasRelacionGastos.flexGastos.Col = 0
                        frmVtasRelacionGastos.flexGastos.Focus()
                    End With
                Case "RELGASTOS.DESCRIPCION AGRUPADOR"
                    With frmVtasRelacionGastos.flexGastos
                        .set_TextMatrix(.Row, 1, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                        'frmPagos.txtFlex = Trim(Flexdet.TextMatrix(Flexdet.Row, 0))
                        .set_TextMatrix(.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        Me.Close()
                        If frmVtasRelacionGastos.ValidaCodigos((frmVtasRelacionGastos.flexGastos.Row)) Then
                            MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            frmVtasRelacionGastos.flexGastos.set_TextMatrix(frmVtasRelacionGastos.flexGastos.Row, 1, "")
                            frmVtasRelacionGastos.flexGastos.set_TextMatrix(frmVtasRelacionGastos.flexGastos.Row, 2, "")
                            frmVtasRelacionGastos.flexGastos.Col = 2
                            frmVtasRelacionGastos.flexGastos.Focus()
                        Else
                            frmVtasRelacionGastos.flexGastos.set_TextMatrix(frmVtasRelacionGastos.flexGastos.Row, 3, "")
                            frmVtasRelacionGastos.flexGastos.set_TextMatrix(frmVtasRelacionGastos.flexGastos.Row, 4, "")
                            If frmVtasRelacionGastos.flexGastos.Row = frmVtasRelacionGastos.flexGastos.Rows - 1 Then
                                frmVtasRelacionGastos.flexGastos.Rows = frmVtasRelacionGastos.flexGastos.Rows + 1
                            End If
                        End If
                        frmVtasRelacionGastos.flexGastos.Col = 3
                        frmVtasRelacionGastos.flexGastos.Focus()
                    End With
                Case "RELGASTOS.DESCRIPCION RUBRO"
                    With frmVtasRelacionGastos.flexGastos
                        .set_TextMatrix(.Row, 3, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                        .set_TextMatrix(.Row, 4, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        Me.Close()
                        If frmVtasRelacionGastos.ValidaCodigos((frmVtasRelacionGastos.flexGastos.Row)) Then
                            MsgBox("No es posible repetir codigos, favor de verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            frmVtasRelacionGastos.flexGastos.set_TextMatrix(frmVtasRelacionGastos.flexGastos.Row, 3, "")
                            frmVtasRelacionGastos.flexGastos.set_TextMatrix(frmVtasRelacionGastos.flexGastos.Row, 4, "")
                            frmVtasRelacionGastos.flexGastos.Col = 4
                            frmVtasRelacionGastos.flexGastos.Focus()
                        Else
                            If frmVtasRelacionGastos.flexGastos.Row = frmVtasRelacionGastos.flexGastos.Rows - 1 Then
                                frmVtasRelacionGastos.flexGastos.Rows = frmVtasRelacionGastos.flexGastos.Rows + 1
                                frmVtasRelacionGastos.flexGastos.Row = frmVtasRelacionGastos.flexGastos.Row + 1
                                frmVtasRelacionGastos.flexGastos.TopRow = frmVtasRelacionGastos.flexGastos.Row
                            Else
                                frmVtasRelacionGastos.flexGastos.Row = frmVtasRelacionGastos.flexGastos.Row + 1
                                If frmVtasRelacionGastos.flexGastos.Row > 6 Then
                                    frmVtasRelacionGastos.flexGastos.TopRow = frmVtasRelacionGastos.flexGastos.Row
                                End If
                            End If
                        End If
                        frmVtasRelacionGastos.flexGastos.Col = 0
                        frmVtasRelacionGastos.flexGastos.Focus()
                    End With
                    'Termina busquedas para la relacion de gastos
                Case "FRMBANCOSPROCESODIARIOREGISTRODEPAGOS.TXTFOLIOEGRESO"
                    With frmBancosProcesoDiarioRegistrodePagos
                        .txtFolioEgreso.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        frmBancosProcesoDiarioRegistrodePagos.bandera = False
                        .LlenaDatos()
                        Me.Close()
                        frmPagos.Hide()
                    End With
                Case "FRMBANCOSPROCESODIARIOREGISTRODEDEPOSITOS.TXTFOLIOINGRESO"
                    With frmBancosProcesoDiarioRegistrodeDepositos
                        .txtFolioIngreso.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        frmBancosProcesoDiarioRegistrodePagos.bandera = False
                        .LlenaDatos()
                        Me.Close()
                        If Not (frmDepositos Is Nothing) Then frmDepositos.Hide()
                    End With
                Case "FRMBANCOSPROCESODIARIOREGISTRODEDEPOSITOS.TXTFOLIORETIRO"
                    With frmBancosProcesoDiarioRegistrodeDepositos
                        .txtFolioRetiro.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        frmBancosProcesoDiarioRegistrodePagos.bandera = False
                        .LlenaDatosRetiros()
                        Me.Close()
                    End With
                Case "FRMBANCOSPROCESODIARIOREGISTRODEOTROSINGRESOS.TXTFOLIOINGRESO"
                    With frmBancosProcesoDiarioRegistrodeOtrosIngresos
                        .txtFolioIngreso.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        frmBancosProcesoDiarioRegistrodePagos.bandera = False
                        frmBancosProcesoDiarioRegistrodeOtrosIngresos.bandera = False
                        .LlenaDatos()
                        Me.Close()
                        frmOtrosIngresos.Hide()
                    End With
                Case "FRMBANCOSPROCESODIARIOCARGOSDIVERSOS.TXTFOLIOEGRESO"
                    With frmBancosProcesoDiarioCargosDiversos
                        .txtFolioEgreso.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        frmBancosProcesoDiarioCargosDiversos.bandera = False
                        .LlenaDatos()
                        Me.Close()
                        frmCargos.Hide()
                    End With
                Case "FRMBANCOSPROCESODIARIOTRASPASOSBANCARIOS.TXTBANCOORIGEN"
                    With frmBancosProcesoDiarioTraspasosBancarios
                        .txtBancoOrigen.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        Me.Close()
                    End With
                Case "FRMBANCOSPROCESODIARIOTRASPASOSBANCARIOS.TXTBANCODESTINO"
                    With frmBancosProcesoDiarioTraspasosBancarios
                        .txtBancoDestino.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        Me.Close()
                    End With
                Case "FRMBANCOSPROCESODIARIOTRASPASOSBANCARIOS.TXTCUENTABANCARIAORIGEN"
                    With frmBancosProcesoDiarioTraspasosBancarios
                        .txtCuentaBancariaOrigen.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 2))
                        Me.Close()
                        .ChecaCuentaOrigen()
                    End With
                Case "FRMBANCOSPROCESODIARIOTRASPASOSBANCARIOS.TXTCUENTABANCARIADESTINO"
                    With frmBancosProcesoDiarioTraspasosBancarios
                        .txtCuentaBancariaDestino.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 2))
                        Me.Close()
                        .ChecaCuentaDestino()
                    End With
                    '            Case "FRMBANCOSPROCESODIARIOREGISTRODEPAGOS.TXTFOLIOPROGRAMACION"
                    '                With frmBancosProcesoDiarioRegistrodePagos
                    '                    .txtFolioProgramacion = Flexdet.TextMatrix(Flexdet.Row, 0)
                    '                    .intNumPartida = Flexdet.TextMatrix(Flexdet.Row, 1)
                    '                    Unload Me
                    '                    .LlenaDatosProgramacion
                    '                    '.txtFolioProgramacion.SetFocus
                    '                End With
                Case "FRMBANCOSPROCESODIARIOTRASPASOSBANCARIOS.TXTFOLIOEGRESO"
                    With frmBancosProcesoDiarioTraspasosBancarios
                        .txtFolioEgreso.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMBANCOSPROCESODIARIOANTICIPOPROVEEDORESACREED.TXTFOLIOEGRESO"
                    With frmBancosProcesoDiarioAnticipoProveedoresAcreed
                        .txtFolioEgreso.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        frmBancosProcesoDiarioAnticipoProveedoresAcreed.bandera = False
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMBANCOSPROCESODIARIOCANCELACIONDEMOVIMIENTOSBANC.TXTFOLIOMOVIMIENTO"
                    With frmBancosProcesoDiarioCancelaciondeMovimientosBanc
                        .txtFolioMovimiento.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        .LlenaDatosMovimientos()
                        Me.Close()
                    End With
                Case "FRMBANCOSPROCESODIARIOCANCELACIONDEMOVIMIENTOSBANC.TXTFOLIOCANCELACION"
                    With frmBancosProcesoDiarioCancelaciondeMovimientosBanc
                        .txtFolioCancelacion.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        .LlenaDatos()
                        Me.Close()
                    End With
                    'Configuración de Caja del PV
                Case "FRMPVCONFIGCAJA.TXTNUMCAJA"
                    With frmPVConfigCaja
                        .txtNumCaja.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMPVCONFIGCAJA.TXTDESCRIPCION"
                    With frmPVConfigCaja
                        .txtNumCaja.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMBANCOSPROCESOMENSUALCONSULTAORIGENAPLICREC.TXTAGRUPADOR"
                    With frmBancosProcesoMensualConsultaOrigenAplicRec
                        .txtAgrupador.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        .dbcAgrupador.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1))
                        Me.Close()
                    End With
                Case "FRMBANCOSPROCESOMENSUALCONSULTAORIGENAPLICREC.TXTRUBRO"
                    With frmBancosProcesoMensualConsultaOrigenAplicRec
                        .txtRubro.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        .dbcRubro.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1))
                        Me.Close()
                    End With
                    '-----------------------------------------------------------------------------
                    '   MOVIMIENTOS DE INVENTARIOS
                Case "FRMINVSALIDAPORVENTA.TXTFOLIO"
                    'With frmInvSalidaPorVenta
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVSALIDAPOROBSEQUIO.TXTFOLIO"
                    'With frmInvSalidaPorObsequio
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVENTRADAPORDEVOLSOBREOBSEQUIO.TXTFOLIO"
                    'With frmInvEntradaPorDevolSobreObsequio
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVENTRADAPORDEVOLSOBREVENTA.TXTFOLIO"
                    'With frmInvEntradaPorDevolSobreVenta
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVENTRADAPORAJUSTE.TXTFOLIO"
                    'With frmInvEntradaporAjuste
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVSALIDAPORAJUSTE.TXTFOLIO"
                    'With frmInvSalidaporAjuste
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                    'SALIDA POR PRESTAMO
                Case "FRMINVSALIDAPORPRESTAMO.TXTFOLIO"
                    'With frmInvSalidaPorPrestamo
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVSALIDAPORPRESTAMO.TXTDETALLE"
                    'With frmInvSalidaPorPrestamo
                    '    With .msgDetallePrestamo
                    '        'Obtener la columna de donde se está ejecutando la consulta
                    '        Columna = .Col
                    '    End With
                    '    If Columna = 0 Then 'Se Busca por código 'Columna de Codigo
                    '        .msgDetallePrestamo.set_TextMatrix(.msgDetallePrestamo.Row, 0, Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '        .LlenarDatosArticulo(Flexdet.get_TextMatrix(Flexdet.Row, 0), "C")
                    '        Me.Close()
                    '        .msgDetallePrestamo.Focus()
                    '        '.txtDetalle.Visible = False
                    '    ElseIf Columna = 1 Then  'Se busca por descripción
                    '        .msgDetallePrestamo.set_TextMatrix(.msgDetallePrestamo.Row, 0, Flexdet.get_TextMatrix(Flexdet.Row, 1))
                    '        .LlenarDatosArticulo(Flexdet.get_TextMatrix(Flexdet.Row, 1), "C")
                    '        Me.Close()
                    '        .msgDetallePrestamo.Focus()
                    '    End If
                    '    .txtDetalle.Visible = False
                    '    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    '    Exit Sub
                    'End With
                Case "FRMINVENTRADAPORDEVOLPORPRESTAMO.TXTFOLIOSALIDA"
                    'With frmInvEntradaPorDevolPorPrestamo
                    '    .txtFolioSalida.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatosSalida()
                    '    Me.Close()
                    'End With
                Case "FRMINVENTRADAPORDEVOLPORPRESTAMO.TXTFOLIO"
                    'With frmInvEntradaPorDevolPorPrestamo
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With

                    'SALIDA POR TRANSFERENCIA

                Case "FRMINVSALIDAPORTRANSFERENCIA.TXTFOLIO"
                    'With frmInvSalidaPorTransferencia
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVSALIDAPORTRANSFERENCIA.TXTDETALLE", "FRMINVSALIDAPORTRANSFERENCIA.MSGDETALLETRANSFERENCIA"
                    'With frmInvSalidaPorTransferencia
                    '    With .msgDetalleTransferencia
                    '        'Obtener la columna de donde se está ejecutando la consulta
                    '        Columna = .Col
                    '    End With
                    '    If Columna = 0 Then 'Se Busca por código 'Columna de Codigo
                    '        .msgDetalleTransferencia.set_TextMatrix(.msgDetalleTransferencia.Row, 0, Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '        VarAux = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                    '        Me.Close()
                    '        .LlenarDatosArticulo(VarAux, "C")
                    '        .msgDetalleTransferencia.Focus()
                    '    ElseIf Columna = 1 Then  'Se busca por descripción
                    '        .msgDetalleTransferencia.set_TextMatrix(.msgDetalleTransferencia.Row, 0, Flexdet.get_TextMatrix(Flexdet.Row, 1))
                    '        VarAux = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '        Me.Close()
                    '        .LlenarDatosArticulo(VarAux, "C")
                    '        .msgDetalleTransferencia.Focus()
                    '    End If
                    '    .txtDetalle.Visible = False
                    '    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    '    Exit Sub
                    'End With
                    'ENTRADA POR TRANSFERENCIA
                Case "FRMINVENTRADAPORTRANSFERENCIA.TXTFOLIO"
                    'With frmInvEntradaPorTransferencia
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVSALIDAAVENDEDORESEXTERNOS.TXTFOLIO"
                    'With frmInvSalidaAVendedoresExternos
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVENTRADAPORDEVOLDEVENDEDORESEXTERNOS.TXTFOLIO"
                    'With frmInvEntradaPorDevoldeVendedoresExternos
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVENTRADAPORCOMPRA.TXTFOLIO"
                    'With frmInventradaPorCompra
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVSALIDAPORDEVOLSOBRECOMPRA.TXTFOLIO"
                    'With frmInvSalidaPorDevolSobreCompra
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVSALIDAPORVENTAAVENDEDORESEXTERNOS.TXTFOLIO"
                    'With frmInvSalidaPorVentaAVendedoresExternos
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMINVENTRADAPORDEVOLSOBREVENTAAVENDEDORESEXTERNOS.TXTFOLIO"
                    'With frmInvEntradaPorDevolsobreVentaAVendedoresExternos
                    '    .txtFolio.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                    '---------------------------------------------------------------------------
                Case "FRMVTASVEENTRADADEMERCANCIA.TXTCODSUCVENDEXTERNO"
                    With frmVtasVESalidadeMercancia
                        .txtCodSucVendExterno.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        .dbcSucursal.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1))
                        Me.Close()
                    End With
                Case "FRMVTASVESALIDADEMERCANCIA.TXTFLEX", "FRMVTASVESALIDADEMERCANCIA.FLEXDETALLE"
                    With frmVtasVESalidadeMercancia
                        .flexDetalle.set_TextMatrix(.flexDetalle.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        If .txtFlex.Visible = True Then
                            .txtFlex.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        End If
                        Me.Close()
                        If .LlenaDatosArticulos() Then
                            .flexDetalle.Col = 3
                            .flexDetalle.Focus()
                        Else
                            .flexDetalle.Text = ""
                            .txtFlex.Text = ""
                        End If
                        Exit Sub
                    End With
                Case "FRMVTASVESALIDADEMERCANCIA.TXTFOLIO"
                    With frmVtasVESalidadeMercancia
                        .txtFolio.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        frmVtasVESalidadeMercancia.bandera = False
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMVTASVESALIDADEMERCANCIA.TXTCODSUCVENDEXTERNO"
                    With frmVtasVESalidadeMercancia
                        .txtCodSucVendExterno.Text = Numerico(Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                        frmVtasVESalidadeMercancia.bandera = False
                        .BuscaVendedorExterno()
                        Me.Close()
                        Exit Sub
                    End With
                Case "FRMVTASVEENTRADADEMERCANCIA.TXTFOLIO"
                    With frmVtasVEEntradadeMercancia
                        .txtFolio.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        frmVtasVEEntradadeMercancia.bandera = False
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMVTASVEENTRADADEMERCANCIA.TXTFOLIOSALIDA"
                    With frmVtasVEEntradadeMercancia
                        .txtFolioSalida.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        frmVtasVEEntradadeMercancia.bandera = False
                        .LlenaDatosFolioSalida()
                        Me.Close()
                    End With
                Case "FRMVTASVELIQUIDACIONVENDEDOREXTERNO.TXTFOLIO"
                    With frmVtasVELiquidacionVendedorExterno
                        .txtFolio.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        .LlenaDatos()
                        Me.Close()
                    End With
                Case "FRMVTASVELIQUIDACIONVENDEDOREXTERNO.TXTFOLIOENTREGA"
                    With frmVtasVELiquidacionVendedorExterno
                        .txtFolioEntrega.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                        Me.Close()
                        .BuscaExistencias()
                    End With
                Case "FRMRPTKARDEXARTICULO.TXTCODARTICULO"
                    With frmRptKardexArticulo
                        .txtCodArticulo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        .LlenaDatosArticulo(CInt(Val(Flexdet.get_TextMatrix(Flexdet.Row, 0))))
                        Me.Close()
                        Me.Dispose()
                    End With

                Case "FRMRPTKARDEXARTICULO.TXTDESCARTICULO"
                    With frmRptKardexArticulo
                        If .ResBusquedaArt = -2 Then
                            .txtCodArticulo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                            .LlenaDatosArticulo(CInt(Val(Flexdet.get_TextMatrix(Flexdet.Row, 0))))
                        Else
                            .txtCodArticulo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                            .LlenaDatosArticulo(CInt(Val(Flexdet.get_TextMatrix(Flexdet.Row, 1))))
                        End If
                        Me.Close()
                        Me.Dispose()
                        Exit Sub
                    End With

                Case "FRMPROGRAMACIONPROMOCIONES.DTPFECHAINICIOJ"
                    'With frmProgramacionPromociones
                    '    .dtpFechaInIcioJ._Value = Flexdet.get_TextMatrix(Flexdet.Row, 3)
                    '    .dtpFechaFinJ._Value = Flexdet.get_TextMatrix(Flexdet.Row, 4)
                    '    .dtpFechaInIcioR._Value = Flexdet.get_TextMatrix(Flexdet.Row, 3)
                    '    .dtpFechaFinR._Value = Flexdet.get_TextMatrix(Flexdet.Row, 4)
                    '    .dtpFechaInIcioV._Value = Flexdet.get_TextMatrix(Flexdet.Row, 3)
                    '    .dtpFechaFinV._Value = Flexdet.get_TextMatrix(Flexdet.Row, 4)
                    '    If Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)) = "JOYERIA" Then
                    '        Grid = 0
                    '    ElseIf Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)) = "RELOJERIA" Then
                    '        Grid = 1
                    '    ElseIf Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)) = "VARIOS" Then
                    '        Grid = 2
                    '    ElseIf Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)) = "X ARTICULO" Then
                    '        Grid = 3
                    '    ElseIf Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0)) = "ARTICULO X PROV" Then
                    '        Grid = 4
                    '        .mProveedor = CShort(ModEstandar.Numerico(Trim(Flexdet.get_TextMatrix(Flexdet.Row, 7))))
                    '        .mrenProv = CShort(ModEstandar.Numerico(Trim(Flexdet.get_TextMatrix(Flexdet.Row, 6))))
                    '    End If
                    '    Me.Close()
                    '    .LlenaDatos(Grid)
                    '    Exit Sub
                    'End With
                Case "FRMPROGRAMACIONPROMOCIONES.DTPFECHAINICIOR"
                    'With frmProgramacionPromociones
                    '    .dtpFechaInIcioJ._Value = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .dtpFechaFinJ._Value = Flexdet.get_TextMatrix(Flexdet.Row, 2)
                    '    .dtpFechaInIcioR._Value = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .dtpFechaFinR._Value = Flexdet.get_TextMatrix(Flexdet.Row, 2)
                    '    .dtpFechaInIcioV._Value = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .dtpFechaFinV._Value = Flexdet.get_TextMatrix(Flexdet.Row, 2)
                    '    Me.Close()
                    '    .LlenaDatos(1)
                    '    Exit Sub
                    'End With
                Case "FRMPROGRAMACIONPROMOCIONES.DTPFECHAINICIOV"
                    'With frmProgramacionPromociones
                    '    .dtpFechaInIcioJ._Value = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .dtpFechaFinJ._Value = Flexdet.get_TextMatrix(Flexdet.Row, 2)
                    '    .dtpFechaInIcioR._Value = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .dtpFechaFinR._Value = Flexdet.get_TextMatrix(Flexdet.Row, 2)
                    '    .dtpFechaInIcioV._Value = Flexdet.get_TextMatrix(Flexdet.Row, 1)
                    '    .dtpFechaFinV._Value = Flexdet.get_TextMatrix(Flexdet.Row, 2)
                    '    Me.Close()
                    '    .LlenaDatos(2)
                    '    Exit Sub
                    'End With
                Case "FRMPROGRAMACIONPROMOCIONES.TXTARTICULO"
                    'With frmProgramacionPromociones
                    '    CodArt = CInt(Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    .LlenaDatosArticulo(CodArt, 1)
                    '    Exit Sub
                    'End With
                Case "FRMPROGRAMACIONPROMOCIONES.MSGJOYERIA"
                    'With frmProgramacionPromociones
                    '    CodArt = CInt(Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    .LlenaDatosArticulo(CodArt, 1)
                    '    Exit Sub
                    'End With
                Case "FRMPROGRAMACIONPROMOCIONES.TXTARTICULOR"
                    'With frmProgramacionPromociones
                    '    CodArt = CInt(Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    .LlenaDatosArticulo(CodArt, 2)
                    '    Exit Sub
                    'End With
                Case "FRMPROGRAMACIONPROMOCIONES.MSGRELOJERIA"
                    'With frmProgramacionPromociones
                    '    CodArt = CInt(Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    .LlenaDatosArticulo(CodArt, 2)
                    '    Exit Sub
                    'End With
                Case "FRMPROGRAMACIONPROMOCIONES.TXTARTICULOV"
                    'With frmProgramacionPromociones
                    '    CodArt = CInt(Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    .LlenaDatosArticulo(CodArt, 3)
                    '    Exit Sub
                    'End With
                Case "FRMPROGRAMACIONPROMOCIONES.MSGVARIOS"
                    'With frmProgramacionPromociones
                    '    CodArt = CInt(Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    Me.Close()
                    '    .LlenaDatosArticulo(CodArt, 3)
                    '    Exit Sub
                    'End With
                Case "FRMPROGRAMACIONPROMOCIONES.MSGXARTICULO"
                    'With frmProgramacionPromociones
                    '    If frmProgramacionPromociones.msgXArticulo.Col = 0 Then
                    '        CodArt = CInt(Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '        .msgXArticulo.set_TextMatrix(.msgXArticulo.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 2)))
                    '        Me.Close()
                    '        .LlenaDatosXArticulo(CodArt)
                    '        Exit Sub
                    '    ElseIf frmProgramacionPromociones.msgXArticulo.Col = 1 Then
                    '        CodArt = CInt(Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '        .msgXArticulo.set_TextMatrix(.msgXArticulo.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 2)))
                    '        Me.Close()
                    '        .LlenaDatosXArticulo(CodArt)
                    '        Exit Sub
                    '    End If
                    'End With
                Case "FRMPROGRAMACIONPROMOCIONES.TXTFLEX"
                    'With frmProgramacionPromociones
                    '    If frmProgramacionPromociones.msgXArticulo.Col = 0 Then
                    '        CodArt = CInt(Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '        .msgXArticulo.set_TextMatrix(.msgXArticulo.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 2)))
                    '        Me.Close()
                    '        .LlenaDatosXArticulo(CodArt)
                    '        Exit Sub
                    '    ElseIf frmProgramacionPromociones.msgXArticulo.Col = 1 Then
                    '        CodArt = CInt(Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '        .msgXArticulo.set_TextMatrix(.msgXArticulo.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 2)))
                    '        Me.Close()
                    '        .LlenaDatosXArticulo(CodArt)
                    '        Exit Sub
                    '    End If
                    'End With

                Case "FRMPROGRAMACIONPROMOCIONES.TXTDETARTXPROV"
                    'With frmProgramacionPromociones
                    '    If frmProgramacionPromociones.msgArtxProv.Col = 0 Then
                    '        CodArt = CInt(ModEstandar.Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '        .msgArtxProv.set_TextMatrix(.msgArtxProv.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 2)))
                    '        Me.Close()
                    '        .LlenaDatosArtxProv(CodArt)
                    '        Exit Sub
                    '    ElseIf frmProgramacionPromociones.msgArtxProv.Col = 1 Then
                    '        CodArt = CInt(Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '        .msgArtxProv.set_TextMatrix(.msgArtxProv.Row, 2, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 2)))
                    '        Me.Close()
                    '        .LlenaDatosArtxProv(CodArt)
                    '        Exit Sub
                    '    End If
                    'End With

                Case "FRMABCRFC.TXTCODIGO"
                    'With frmABCRFC
                    '    .txtCodigo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMABCRFC.TXTRFC"
                    'With frmABCRFC
                    '    .txtCodigo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 2)
                    '    .LlenaDatos()
                    '    Me.Close()
                    'End With
                Case "FRMPVREGCHEQUE.TXTCHEQUE"
                    'With frmPVRegCheque
                    '    With .msgCheque
                    '        .set_TextMatrix(.Row, 6, Numerico(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '        .set_TextMatrix(.Row, 0, Trim(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '        Me.Close()
                    '        Exit Sub
                    '    End With
                    'End With
                Case "FRMFACTANALISISVENTAS.TXTFOLIOFACTURA"
                    With frmFactAnalisisVentas
                        .txtFolioFactura.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        Me.Close()
                        .LlenaDatos()
                        Exit Sub
                    End With
                Case "FRMIMPRESIONETIQUETAS.MSGETIQUETAS", "FRMIMPRESIONETIQUETAS.TXTDETALLE"
                    'With frmImpresionEtiquetas
                    '    If .msgEtiquetas.Col = 0 Then
                     '        VarAux = CInt(Val(Flexdet.get_TextMatrix(Flexdet.Row, 0)))
                    '    ElseIf .msgEtiquetas.Col = 1 Then
                     '        VarAux = CInt(Val(Flexdet.get_TextMatrix(Flexdet.Row, 1)))
                    '    End If
                    '    Me.Close()
                    '    .LlenarDatosArticulo(CInt(VarAux), "C")
                    '    Exit Sub
                    'End With
                Case "FRMIMPRESIONETIQUETAS.TXTORDENCOMPRA"
                    'With frmImpresionEtiquetas
                    '    .txtOrdenCompra.Text = Trim(Flexdet.get_TextMatrix(Flexdet.Row, 0))
                    '    Me.Close()
                    '    .MostrarArticulo(0, Trim(.txtOrdenCompra.Text), "O")
                    '    .txtOrdenCompra.Tag = Trim(.txtOrdenCompra.Text)
                    '    Exit Sub
                    'End With
                Case "FRMVERIFICADORPRECIOS.TXTCODARTICULO", "FRMVERIFICADORPRECIOS.TXTDESCARTICULO"
                    With frmVerificadorPrecios
                        .txtCodArticulo.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                        .LlenaDatos(CInt(Val(Flexdet.get_TextMatrix(Flexdet.Row, 0))))
                        .LlenaDatosPromocion(CInt(Val(Flexdet.get_TextMatrix(Flexdet.Row, 0))))
                        Me.Close()
                        Exit Sub
                    End With
                Case "FRMPVREGNOTASCRED.TXTNOTACREDITO"
                    'With frmPVRegNotasCred
                    '    .txtNotaCredito.Text = Flexdet.get_TextMatrix(Flexdet.Row, 0)
                    '    Me.Close()
                    '    .LlenaDatoVale()
                    '    Exit Sub
                    'End With

                    '''03MAR2008 - MAVF
                Case "FRMVTASREPORTEDECUENTASPORCOBRAR.TXTNOMBRE"
                    With frmVtasReportedeCuentasporCobrar
                        .gFueraChange = True
                        .txtCodCliente.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .gFueraChange = False
                        .LlenaDatos()
                        Me.Close()
                    End With

                Case "FRMVTASRPTVENTASSALIDADEMERCANCIAPORCLIENTE.TXTNOMBRE"
                    With frmVtasRPTVentasSalidadeMercanciaPorCliente
                        .gBlnFueraChange = True
                        .txtCodCliente.Text = Flexdet.get_TextMatrix(Me.Flexdet.Row, 0)
                        .gBlnFueraChange = False
                        .LlenaDatos()
                        Me.Close()
                    End With

                Case Else
                    Me.Close()
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    Exit Sub
            End Select
            'System.Windows.Forms.SendKeys.Send("{ENTER}")
        End With
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub
Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MostrarError("Ha ocurrido un error")
    End Sub

    Private Sub FrmConsultas_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.Close()
            If UCase(Trim(System.Windows.Forms.Form.ActiveForm.Name)) = UCase(Trim("frmVtasRelacionGastos")) Then
                frmVtasRelacionGastos.chkFueraEnter.CheckState = System.Windows.Forms.CheckState.Unchecked
            ElseIf UCase(Trim(System.Windows.Forms.Form.ActiveForm.Name)) = UCase(Trim("frmVtasEstadodeResultados")) Then
                frmVtasEstadodeResultados.chkFueraEnter.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
        End If
    End Sub

    Public Sub FrmConsultas_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        KeyPreview = True
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        bandera = True
        'System.Windows.Forms.SendKeys.Send("{RIGHT}")
    End Sub

    Private Sub FrmConsultas_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        IsNothing(Me)
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmConsultas))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Flexdet = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        CType(Me.Flexdet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Flexdet
        '
        Me.Flexdet.DataSource = Nothing
        Me.Flexdet.Location = New System.Drawing.Point(12, 12)
        Me.Flexdet.Name = "Flexdet"
        Me.Flexdet.OcxState = CType(resources.GetObject("Flexdet.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Flexdet.Size = New System.Drawing.Size(648, 281)
        Me.Flexdet.TabIndex = 9
        '
        'FrmConsultas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(672, 303)
        Me.Controls.Add(Me.Flexdet)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(196, 148)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmConsultas"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = " "
        CType(Me.Flexdet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

End Class