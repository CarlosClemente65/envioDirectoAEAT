﻿Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports Microsoft.VisualBasic

Public Class Form1
    Dim lista_envio As New ArrayList
    Dim fich_respuesta As String = ""
    Dim salir As Boolean = False
    Dim viene_de_ds As Boolean = False
    Dim Caducado As String = "Caducado "
    Dim certificado_fich = ""
    Dim certificado_passw = ""


    Private Sub carga_certificados()
        Dim store As X509Store = New X509Store(StoreName.My, StoreLocation.CurrentUser)
        Dim fecha As Date
        Dim serialnumber, cadena As String

        store.Open(OpenFlags.ReadOnly)

        ' Dim certificate As X509Certificate = store.Certificates.Item(0) ' certificates(0)
        Dim certificado As X509Certificate2
        Dim dni As String, repre_dni As String, repre_name As String, AuxCaducado As String

        For x = 0 To store.Certificates.Count - 1
            AuxCaducado = ""
            dni = ""
            repre_dni = ""
            repre_name = ""
            certificado = store.Certificates.Item(x)

            cadena = saca_nombre(certificado.Subject, dni, repre_dni, repre_name)
            If Trim(repre_name) <> "" Then
                cadena = Trim(repre_name)
            End If

            serialnumber = certificado.GetSerialNumberString()

            If cadena <> "" Then
                fecha = Convert.ToDateTime(certificado.GetExpirationDateString())
                If fecha < DateTime.Now Then AuxCaducado = Caducado

                ListBox1.Items.Add(AuxCaducado & certificado.GetExpirationDateString() & " " & cadena)
                ListBox2.Items.Add(serialnumber & " " & x.ToString())
            End If

        Next

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' salir
        Close()

    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If salir Then Close()
    End Sub

    Private Sub Form1_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        If viene_de_ds Then
            If Not File.Exists(fich_respuesta) Then
                File.WriteAllText(fich_respuesta, "MENSAJE = El cliente ha seleccionado salir o no hay certificado seleccionado")
            End If
        End If

    End Sub

    Private Function saca_indice_certificado(dscertificado As String, Caducado As String)
        Dim indice As Integer = -1
        Dim x, y As Integer
        Dim cadena, aux As String

        For x = 0 To ListBox1.Items.Count - 1
            cadena = ListBox1.Items(x)

            aux = Microsoft.VisualBasic.Left(cadena, Len(Caducado))

            y = InStr(1, cadena, dscertificado)
            If (InStr(1, cadena, dscertificado) <> 0) And (aux <> Caducado) Then
                indice = x
                Exit For
            End If
        Next

        Return indice
    End Function
    Private Function saca_indice_certificado_serie(dscertificado)
        Dim indice As Integer = -1
        Dim x As Integer
        Dim cadena As String

        dscertificado = dscertificado.ToUpper()

        For x = 0 To ListBox2.Items.Count - 1
            Dim argumentos As String()
            cadena = ListBox2.Items(x).ToUpper()
            argumentos = cadena.Split() ' viene con dos datos el segundo es el numero de indice del certificado que se ha guardado antes

            If argumentos(0) = dscertificado Then
                indice = x
                Exit For
            End If

        Next

        Return indice
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' ds123456 0 C:\vb_net\Ds_SII\modelo_180.txt C:\vb_net\Ds_SII\respuesta.res
        ' ds123456 0 C:\vb_net\Ds_SII\modelo_180.txt C:\vb_net\Ds_SII\respuesta.res  "INFORMATICA PARA LA PEQUEña"
        ' ds123456 0 C:\vb_net\Ds_SII\modelo_180.txt C:\vb_net\Ds_SII\respuesta.res NO 047fa4b5c25d36a159391887268cfc38
        ' ds123457 0 C:\AEAT\111\111_2021.txt C:\AEAT\111\111_2021.res  "c:\certificados\Carlos Clemente (28.02.25).pfx" 05196375P

        Dim arguments As String() = Environment.GetCommandLineArgs()
        Dim dsclave As String = ""
        Dim tipo, dscertificado As String
        Dim indice As Integer = -1

        Me.Height = 500
        Me.Width = 900
        Me.Text = "Envio modelos AEAT - Diagram Software S.L. v1.2 "


        ' ver si coger certificados de almacen o viene el certificado y password en parametros
        If arguments.Count > 1 Then
            dsclave = arguments(1)
            Select Case dsclave
                Case "ds123456"  ' proceso con almacen de certificados
                    ' ejemplo ds123456 0 C:\AEAT\111\111_2021.txt C:\AEAT\111\111_2021.res

                    carga_certificados()
                    If arguments.Count >= 5 Then
                        tipo = arguments(2)
                        ComboBox1.SelectedIndex = 0
                        ComboBox1.Enabled = False
                        TextBox1.Text = arguments(3) ' fichero a enviar
                        TextBox1.Enabled = False
                        Button2.Enabled = False
                        fich_respuesta = arguments(4) ' // xml recibe respuesta web service
                        viene_de_ds = True

                        If File.Exists(fich_respuesta) Then File.Delete(fich_respuesta)

                        If arguments.Count = 6 Then
                            ' viene con nombre de certificado ver el indice al que pertence dicho certificado
                            dscertificado = arguments(5)
                            indice = saca_indice_certificado(dscertificado, Caducado)
                        End If

                        If arguments.Count = 7 Then
                            ' viene con nombre de certificado ver el indice al que pertence dicho certificado
                            dscertificado = arguments(6)
                            indice = saca_indice_certificado_serie(dscertificado)  ' Numero de serie
                        End If

                        If indice <> -1 Then
                            ListBox1.SelectedIndex = indice
                            Button3.PerformClick()  ' lanzamos el proceso y si no presentamos pantalla para enviarlo cuando se pulse el boton enviar.
                        End If

                    Else
                        End
                    End If

                Case "ds123457" ' proceso con fichero de certificados y password
                    ' ejemplo ds123457 0 C:\AEAT\111\111_2021.txt C:\AEAT\111\111_2021.res  "c:\certificados\Carlos Clemente (28.02.25).pfx" 05196375P
                    If arguments.Count = 7 Then
                        tipo = arguments(2)
                        TextBox1.Text = arguments(3) ' fichero a enviar
                        fich_respuesta = arguments(4) ' // xml recibe respuesta web service
                        If File.Exists(fich_respuesta) Then File.Delete(fich_respuesta)
                        certificado_fich = arguments(5) ' fichero certificado
                        certificado_passw = arguments(6) ' password certificado
                        viene_de_ds = True
                        envio_peticion()

                    Else
                        End  ' terminamos

                    End If

                Case Else
                    End ' ha entrado con clave incorrecta
            End Select


        Else
            ' no ha introducido parametros pedir clave
            If File.Exists(fich_respuesta) Then File.Delete(fich_respuesta)
            dsclave = InputBox("Introducir contraseña", "")
            viene_de_ds = False
            If dsclave <> "ds123456" Then End ' terminamos
            carga_certificados()
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Function saca_codificacion_modelo(fichero As String) As String
        Dim cadena, valor As String
        Dim lista As New ArrayList
        Dim sw As Integer = 0
        Dim arra As Array

        valor = ""
        Using sr As New StreamReader(fichero)
            Dim line As String
            ' Read and display lines from the file until the end of
            ' the file is reached.
            Do
                line = sr.ReadLine()
                If Not (line Is Nothing) Then
                    lista.Add(line)
                End If
            Loop Until line Is Nothing
        End Using


        For x = 0 To lista.Count - 1
            cadena = lista(x)

            If Trim(cadena) <> "" Then
                If Trim(cadena) = "[url]" Then
                    sw = 1
                    Continue For
                End If

                If Trim(cadena) = "[cabecera]" Then
                    sw = 2
                    Continue For
                End If

                If Trim(cadena) = "[body]" Then
                    sw = 3
                    Continue For
                End If

                If Trim(cadena) = "[respuesta]" Then
                    sw = 4
                    Continue For
                End If
                '

                If sw = 2 Then
                    If Mid(Trim(cadena), 1, 12) = "CODIFICACION" Then
                        arra = cadena.Split("=")
                        If arra.Length > 1 Then
                            valor = Trim(arra(1))
                        End If
                        Exit For
                    End If
                End If
                
            End If
        Next

        If valor = "" Then valor = "utf-8"
        Return valor.ToUpper


    End Function


    Private Sub carga_datos_modelo(fichero As String, codificacion As String, ByRef cabecera As ArrayList, ByRef body As ArrayList, ByRef respuesta As ArrayList, ByRef url As String)
        Dim cadena As String
        Dim lista As New ArrayList
        Dim sw As Integer = 0

        Using sr As New StreamReader(fichero, Encoding.GetEncoding(codificacion))
            Dim line As String
            ' Read and display lines from the file until the end of
            ' the file is reached.
            Do
                line = sr.ReadLine()
                If Not (line Is Nothing) Then
                    lista.Add(line)
                End If
            Loop Until line Is Nothing
        End Using


        For x = 0 To lista.Count - 1
            cadena = lista(x)

            If Trim(cadena) <> "" Then
                If Trim(cadena) = "[url]" Then
                    sw = 1
                    Continue For
                End If

                If Trim(cadena) = "[cabecera]" Then
                    sw = 2
                    Continue For
                End If

                If Trim(cadena) = "[body]" Then
                    sw = 3
                    Continue For
                End If

                If Trim(cadena) = "[respuesta]" Then
                    sw = 4
                    Continue For
                End If
                '
                If sw = 1 Then
                    url = cadena
                End If
                If sw = 2 Then
                    cabecera.Add(cadena)
                End If
                If sw = 3 Then
                    body.Add(cadena)
                End If
                If sw = 4 Then
                    respuesta.Add(cadena)
                End If
            End If
        Next

    End Sub


    Private Sub saca_header(cadena As String, ByRef titulo As String, ByRef valor As String)
        ' Dim arra As Array
        ' titulo = ""
        ' valor = ""
        ' Try
        'arra = cadena.Split("=")
        'titulo = Trim(arra(0))
        'valor = Trim(arra(1))
        'Catch ex As Exception
        '
        '       End Try

        Dim FirstCharacter As Integer
        titulo = ""
        valor = ""
        Try
            FirstCharacter = cadena.IndexOf("=")  ' Devuelve el índice de base cero de la primera aparición

            titulo = Trim(Mid(cadena, 1, FirstCharacter))
            valor = Trim(Mid(cadena, FirstCharacter + 2))
        Catch ex As Exception

        End Try

    End Sub
    Private Function quita_raros(aux As String)
        aux = aux.Replace("á", "a").Replace("é", "e").Replace("í", "i").Replace("ó", "o").Replace("ú", "u").Replace("º", ".").Replace("ª", ".").Replace("ñ", "¤")

        aux = aux.Replace("Á", "A").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Ú", "U").Replace("ñ", "¤")
        Return aux
    End Function

    Private Sub envio_peticion() ' lo envia como si fuera un formulario con el submit 
        Dim indice_certificado, x As Integer
        Dim argumentos As String()
        Dim Cabecera As New ArrayList, body As New ArrayList, respuesta As New ArrayList
        Dim titulo, valor, aux, codifica, cadena, estado As String
        Dim url As String = ""
        Dim valido As Boolean = False

        Try
            System.Net.ServicePointManager.SecurityProtocol = 3072 ' Tls12	3072  - Tls11	768  para que no de error de que no puede crear un canal seguro. Con el framework 4.7.2 ya no da error si se quita esta linea
            codifica = saca_codificacion_modelo(TextBox1.Text) ' UTF-8 , ISO8859-1 (ascii extendido 256 bits o ansi)
            carga_datos_modelo(TextBox1.Text, codifica, Cabecera, body, respuesta, url)

            Dim request As HttpWebRequest = WebRequest.Create(url)
            Dim certificate As X509Certificate2  ' X509Certificate deberia ser X509Certificate2 para proxima compilacion, ya se ha cambiado falta compilarlo

            If certificado_fich = "" Then
                cadena = ListBox2.Items(ListBox1.SelectedIndex)
                argumentos = cadena.Split()
                indice_certificado = Integer.Parse(argumentos(1))
                certificate = asigna_certificado_almacen(indice_certificado)
            Else
                certificate = carga_certificado(certificado_fich, certificado_passw)
            End If

            request.ClientCertificates.Add(certificate)

            Dim data As StringBuilder
            data = New StringBuilder()
            ' preparar json para enviar
            data.Append("{")
            For x = 0 To Cabecera.Count - 1
                titulo = ""
                valor = ""
                saca_header(Cabecera(x), titulo, valor)

                If titulo = "F01" Then
                    '   valor = valor.Replace(Chr(34), "'") ' si viene un xml se sustituye las comillas dobles en simples.
                    valor = valor.Replace(Chr(34), "\" & Chr(34))
                End If


                If (x = 0) Then
                    data.Append(Chr(34) & titulo & Chr(34) & ":" & Chr(34) & valor & Chr(34))
                Else
                    data.Append("," & Chr(34) & titulo & Chr(34) & ":" & Chr(34) & valor & Chr(34))
                End If
            Next
            data.Append("}")

            '  Dim basura As String = data.ToString()


            Dim basura As String = data.ToString()
            Dim byteArray As Byte()

            ' byteArray = Encoding.Default.GetBytes(data.ToString()) ' ansi
            byteArray = Encoding.UTF8.GetBytes(data.ToString())
            request.Method = "POST"
            request.ContentType = "application/json;charset=UTF-8"
            request.ContentLength = byteArray.Length ' Set the ContentLength property of the WebRequest.

            Dim dataStream As Stream = request.GetRequestStream() ' Get the request stream. 
            dataStream.Write(byteArray, 0, byteArray.Length) ' Write the data to the request stream.
            dataStream.Close() ' Close the Stream object.

            Dim response As WebResponse = request.GetResponse()  ' Get the response. Aqui hace la peticion
            Dim stream As System.IO.Stream = response.GetResponseStream()
            'Dim reader = New StreamReader(stream, System.Text.Encoding.Default, False, 512)
            Dim reader = New StreamReader(stream)

            ' Read from the stream object using the reader, put the contents in a string
            Dim contents As String = reader.ReadToEnd()

            estado = (CType(response, HttpWebResponse).StatusDescription)
            If estado = "OK" Then
                TextBox2.Text = ""
                Dim FirstCharacter As Integer

                'contents = My.Computer.FileSystem.ReadAllText("C:\AEAT\390\2021\001000\errores_varios_errores.html")  ' para pruebas QUITAR DESPUES

                'contents = My.Computer.FileSystem.ReadAllText("C:\AEAT\390\2021\001000\errores_unsolo_error.html")  ' para pruebas QUITAR DESPUES
                'contents = My.Computer.FileSystem.ReadAllText("C:\AEAT\390\2021\001000\errores_casilla.html")

                Dim inicio As Integer
                Dim car As String
                Dim busca As String = ""
                Dim kk As Integer = -1

                For x = 0 To respuesta.Count - 1 ' busca la variables de la respuesta y las rellena con el valor
                    valor = ""
                    aux = Chr(34) & respuesta(x) & Chr(34) & ":" & Chr(34)
                    FirstCharacter = contents.IndexOf(aux) ' distingue entre mayusculas y minisculas
                    If FirstCharacter <> -1 Then
                        inicio = FirstCharacter + Len(aux) + 1
                        For bucle = inicio To Len(contents)
                            car = Mid(contents, bucle, 1)
                            If car = Chr(34) Then
                                valor = Mid(contents, inicio, bucle - inicio).Trim
                                valido = True
                                Exit For
                            End If
                        Next
                    Else
                        ' vemos si es una variable de error
                        aux = Chr(34) & respuesta(x) ' ejemplo "E00
                        FirstCharacter = contents.IndexOf(aux) ' distingue entre mayusculas y minisculas
                        If FirstCharacter <> -1 Then
                            aux = ":["
                            'Dim delimitadores() As String = {"["}
                            Dim delimitadores() As String = {aux}
                            aux = Chr(34) & ","
                            Dim delimitadores1() As String = {aux}
                            Dim vectoraux() As String
                            Dim vectoraux1() As String
                            Dim Final As String
                            Dim itemaux As String
                            inicio = FirstCharacter + Len(aux) + 1
                            Final = Mid(contents, inicio, Len(contents))

                            vectoraux = Final.Split(delimitadores, StringSplitOptions.None) 'en este proceso se graba todos los errores E00, E01 ect...

                            Dim resError As String = ""

                            For Each item As String In vectoraux
                                If kk <> -1 Then
                                    vectoraux1 = item.Split(delimitadores1, StringSplitOptions.None)
                                    For Each item1 As String In vectoraux1
                                        itemaux = item1
                                        If Mid(itemaux, Len(itemaux), 1) = "," Then
                                            itemaux = Mid(itemaux, 1, Len(itemaux) - 1)
                                        End If
                                        itemaux = itemaux.Replace("]}}", "")
                                        If Mid(itemaux, Len(itemaux), 1) <> Chr(34) Then
                                            itemaux = itemaux & Chr(34)
                                        End If

                                        resError = "E" & kk.ToString().PadLeft(2, "0"c)
                                        TextBox2.Text = TextBox2.Text & resError & " = " & itemaux & vbCrLf
                                    Next
                                End If
                                kk = kk + 1
                            Next
                        End If
                    End If

                    If kk = -1 Then ' no ha encontrado errores
                        TextBox2.Text = TextBox2.Text & respuesta(x) & " = " & valor & vbCrLf
                    End If

                Next
            End If

            If Not valido Then
                Dim ruta As String
                ruta = Path.GetDirectoryName(fich_respuesta)
                If ruta = "" Then
                    ruta = My.Application.Info.DirectoryPath
                End If
                aux = quita_raros(contents)
                File.WriteAllText(ruta & "\errores.html", aux, Encoding.Default)
            End If

        Catch ex As Exception
            TextBox2.Text = "MENSAJE = Proceso cancelado o error en envio." & ex.Message
            ' MsgBox("Error .... " & ex.Message)
        End Try

        Try
            aux = TextBox2.Text
            aux = quita_raros(aux)
            If fich_respuesta <> "" Then File.WriteAllText(fich_respuesta, aux, Encoding.Default)
        Catch ex As Exception

        End Try

        If viene_de_ds Then Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' validar entradas

        If ComboBox1.SelectedIndex = -1 Then
            MsgBox("Seleccione el tipo de modelo a enviar")
            Exit Sub
        End If

        If ListBox1.SelectedIndex = -1 Then
            MsgBox("Seleccione un certificado válido")
            Exit Sub
        End If

        Dim fichero As String
        fichero = Trim(TextBox1.Text)
        If fichero = "" Or Not File.Exists(fichero) Then
            MsgBox("Seleccione un archivo válido")
            Exit Sub
        End If

        TextBox2.Text = "" '
        TextBox2.Refresh()

        ' enviamos el fichero, con 0 del almacen de certificados
        envio_peticion()

    End Sub


End Class