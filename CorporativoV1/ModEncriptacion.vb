'**********************************************************************************************************************'
'*PROGRAMA: MODULO DE ENCRIPTACIÓN JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Public Module ModEncriptacion
    '1314282103
    'dias de nacimiento Erica, Carla,Ezequiel,Fernando,Maricruz

    Public Function Encriptar(ByRef text As String) As String
        'On Error GoTo Errores
        Dim Encriptar1 As String = ""
        Try
            Dim ENCR As String
            Dim i As Integer
            ENCR = ""

            If Len(Trim(text)) > 0 Then
                For i = 1 To Len(text)
                    If i = 1 Then
                        ENCR = CStr((Asc(Mid(text, i, 1))) + 1314282103)
                    Else
                        ENCR = ENCR + ";" + CStr((Asc(Mid(text, i, 1))) + 1314282103)
                    End If
                Next i
                Encriptar1 = ENCR

            Else
                Encriptar1 = 1314282103
            End If
            'Errores:
            If Err.Number <> 0 Then ModErrores.Errores()
        Catch ex As Exception
        End Try
        Return Encriptar1
    End Function

    Public Function Desencriptar(ByRef text As String) As String
        'On Error GoTo Errores
        Dim Desenc1 As String = ""
        Try
            Dim Desenc As String
            Dim i As Integer
            Desenc = ""
            Desencriptar = ""

            If text = "1314282103" Then
                Desencriptar = ""
                Exit Function
            End If

            If Len(Trim(text)) > 0 Then
                For i = 1 To Len(text)
                    If Mid(text, i, 1) <> ";" Then
                        Desenc = Desenc & CStr(Mid(text, i, 1))
                    Else
                        Desencriptar = Desencriptar & Chr(CDbl(Desenc) - 1314282103)
                        Desenc = ""
                    End If
                Next i
                Desenc1 = Desencriptar & Chr(CDbl(Desenc) - 1314282103)
            End If
            'Errores:
            If Err.Number <> 0 Then ModErrores.Errores()
        Catch ex As Exception
        End Try
        Return Desenc1
    End Function

End Module