'**********************************************************************************************************************'
'*PROGRAMA: VERIFICAR CONEXION JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports System.IO
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility


Public Class frmVerificarConexion

    Inherits System.Windows.Forms.Form

    Public Sub Cerrar()
        'Unload Me
        Me.Close()
    End Sub

    Public Sub Guardar()
        'On Error GoTo Error   
        Try
            Select Case MsgBox("¿Desea Guardar la Información?", vbYesNo + vbQuestion, "")
                Case vbNo
                    'Unload Me
                    Me.Close()
                    End
            End Select

            If Trim(TxtNomServidor.Text) = "" Then
                MsgBox("Información Incompleta... Proporcione un Nombre del Servidor SQL", vbInformation, "MENSAJE")
                Exit Sub
            End If
            If Trim(TxtBDPrincipal.Text) = "" Then
                MsgBox("Información Incompleta... Proporcione un Nombre de la Base de Datos Principal", MsgBoxStyle.Information, "MENSAJE")
                Exit Sub
            End If

            If ModConexion.Abrir(TxtNomServidor.Text, TxtBDPrincipal.Text) = True Then
                NombreServidor = Trim(Me.TxtNomServidor.Text)
                NombreBaseDatos = Trim(Me.TxtBDPrincipal.Text)

                ArchivoTxt = CreateObject("Scripting.FileSystemObject")
                F = ArchivoTxt.CreateTextFile(rutaArchivoTxt, True)
                'F = ArchivoTxt.CreateTextFile("C:~\\CorporativoV1\\CorporativoV1" & "\\Sistema\\CJoyeria.Txt", True)
                F.Write(TxtNomServidor.Text & vbNewLine & TxtBDPrincipal.Text)
                'banderaConexion = True

                F.Close()

                MsgBox("El archivo principal de inicio del sistema fué creado exitosamente !", vbInformation, "Mensaje")

                Me.Hide()
                Dim frmAccceso As FrmAcceso = New FrmAcceso()
                frmAccceso.Show()


                If Dir(My.Application.Info.DirectoryPath & "\\Sistema", FileAttribute.Directory) = "" Then
                Else
                    ChDir(My.Application.Info.DirectoryPath & "\\Sistema")
                End If
            Else
                TxtNomServidor.Text = ""
                TxtBDPrincipal.Text = ""
            End If

        Catch ex As Exception
            MessageBox.Show("MENSAJE:" + ex.Message)
            'Error
            If Err.Number <> 0 Then
                ModErrores.Errores()
            End If
        End Try


    End Sub




    Private Sub Form_Unload(Cancel As Integer)
        'frmVerificarConexion = Nothing
        'Me.Frame1 = Nothing
    End Sub



    Private Sub frmVerificarConexion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Icono(Me, MDIMenuPrincipalCorpo) 
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
    End Sub

    Private Sub TxtBDPrincipal_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtBDPrincipal.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 0
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
    End Sub

    Private Sub TxtBDPrincipal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtBDPrincipal.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = 13 And TxtBDPrincipal.Text <> "" Then Guardar()
        If KeyCode = 27 Then TxtNomServidor.Focus()
    End Sub

    Private Sub TxtNomServidor_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtNomServidor.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 0
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii)
    End Sub

    Private Sub TxtNomServidor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtNomServidor.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = 27 Then
            If MsgBox("¿Desea Abandonar la Conexión?", vbYesNo + vbQuestion, "JoyeriaMR") <> vbNo Then
                '     Unload Me
                'Unload MenuPrincipal
                'Me.MDIMenuPrincipalCorpo()
                Me.Close()
                End
            End If
        End If
        If KeyCode = 13 And TxtNomServidor.Text <> "" Then TxtBDPrincipal.Focus()
    End Sub







End Class