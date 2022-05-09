Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Module Module1
    Public Function saca_entre_etiquetas(dato As String, eti_desde As String, eti_hasta As String) As String
        Dim cadena As String = ""
        Dim aux As String
        Dim x As Integer = 0
        Dim desde As Integer


        x = dato.IndexOf(eti_desde)
        If x <> -1 Then
            desde = x + Len(eti_desde) + 1
            aux = Mid(dato, desde)

            If eti_hasta = "" Then ' significa que coge desde eti_desde hasta el final de la cadena
                cadena = aux
            Else
                x = aux.IndexOf(eti_hasta)

                If x <> -1 Then ' ha encontrado etiq_hasta y coge el valor entre las dos etiquetas
                    cadena = Mid(aux, 1, x)
                Else
                    cadena = "" ' no encuentra y retornamos nada.
                End If

            End If
        End If

        Return cadena

    End Function

    Public Function busca_CN(cadena As String) As String
        Dim aux As String
        Dim nombre As String = ""
        Dim delimitadores() As String = {"CN="}   ' 
        Dim vectoraux() As String
        Dim subvector() As String
        vectoraux = cadena.Split(delimitadores, StringSplitOptions.None)
        If vectoraux.Count = 2 Then
            aux = vectoraux(1)

            If aux.Substring(0, 1) = """" Then  ' puede empezar por " porque el nombre tiene espacios y lo codifican entre comillas
                subvector = aux.Split(New Char() {""""c})
                nombre = subvector(1)
            Else
                subvector = aux.Split(New Char() {","c})
                nombre = subvector(0)
            End If


        End If

        Return nombre


    End Function

    Public Function busca_A25497(cadena As String) As String
        Dim aux As String
        Dim dni As String = ""
        Dim delimitadores() As String = {"2.5.4.97="}   ' 
        Dim vectoraux() As String
        Dim subvector() As String
        Dim subvector1() As String
        vectoraux = cadena.Split(delimitadores, StringSplitOptions.None)
        If vectoraux.Count = 2 Then
            aux = vectoraux(1)
            subvector = aux.Split(New Char() {","c})
            subvector1 = subvector(0).Split(New Char() {"-"c})
            If subvector1.Count = 2 Then
                dni = subvector1(1)
            Else
                dni = subvector1(0)
            End If
        End If

        Return dni


    End Function

    Public Function busca_SerialNumber(cadena As String) As String
        Dim aux As String
        Dim nombre As String = ""
        Dim delimitadores() As String = {"SERIALNUMBER="}   ' 
        Dim vectoraux() As String

        Dim subvector() As String
        Dim subvector1() As String
        vectoraux = cadena.Split(delimitadores, StringSplitOptions.None)
        If vectoraux.Count = 2 Then
            aux = vectoraux(1)
            subvector = aux.Split(New Char() {","c})
            subvector1 = subvector(0).Split(New Char() {"-"c})
            If subvector1.Count = 2 Then
                nombre = subvector1(1)
            Else
                nombre = subvector1(0)
            End If


        End If

        Return nombre


    End Function
    Public Function saca_nombre(cadena As String, ByRef dni As String, ByRef repre_dni As String, ByRef repre_name As String) As String

        Dim O As String
        Dim CN As String
        Dim A25497 As String
        Dim nombre As String = ""
        Dim SERIALNUMBER As String

        CN = busca_CN(cadena)
        O = busca_O(cadena)
        A25497 = busca_A25497(cadena) ' CIF de la sociedad
        SERIALNUMBER = busca_SerialNumber(cadena)

        If (A25497 <> "") Then  ' Juridica
            If O <> "" Then
                nombre = O ' Razón social del titular del certificado
                dni = A25497 ' CIF de la sociedad
                ' representante de la sociedad.
                repre_dni = SERIALNUMBER
                repre_name = saca_SNyG(cadena) ' apellidos y nombre fisico
            End If

        ElseIf SERIALNUMBER <> "" Then
            If InStr(CN, "FIRMA") = 0 Then
                nombre = saca_SNyG(cadena)
                If nombre = "" Then
                    nombre = CN
                End If
                dni = SERIALNUMBER
            End If
        End If


        Return nombre

    End Function
    Public Function busca_O(cadena As String) As String
        Dim aux As String
        Dim nombre As String = ""
        Dim delimitadores() As String = {"O="}   ' juridica
        Dim vectoraux() As String
        Dim subvector() As String
        vectoraux = cadena.Split(delimitadores, StringSplitOptions.None)
        If vectoraux.Count = 2 Then
            aux = vectoraux(1)

            If aux.Substring(0, 1) = """" Then ' puede empezar por " porque el nombre tiene espacios y lo codifican entre comillas
                subvector = aux.Split(New Char() {""""c})
                nombre = subvector(1)
            Else
                subvector = aux.Split(New Char() {","c})
                nombre = subvector(0)
            End If



        End If

        Return nombre


    End Function
    Public Function saca_SNyG(subjet As String) As String
        Dim aux As String
        Dim x As Integer
        Dim SN As String = ""
        Dim G As String = ""
        Dim delimitadores() As String = {"SN="}   ' 
        Dim delimitadores1() As String = {"G="}

        Dim vectoraux() As String
        Dim subvector() As String
        Dim repre_name = ""
        vectoraux = subjet.Split(delimitadores, StringSplitOptions.None)
        If vectoraux.Count = 2 Then
            aux = vectoraux(1)
            x = aux.IndexOf(",")
            If x <> -1 Then
                SN = Trim(Mid(aux, 1, x))
                subvector = subjet.Split(delimitadores1, StringSplitOptions.None)
                If subvector.Count = 2 Then
                    aux = subvector(1)
                    x = aux.IndexOf(",")
                    If x <> -1 Then
                        G = Trim(Mid(aux, 1, x))
                    End If
                End If
                repre_name = SN & " " & G
            End If
        End If
        Return repre_name
    End Function

    Public Sub saca_representante(subjet As String, ByRef repre_name As String, ByRef repre_cif As String) ' old
        Dim aux, cadena As String
        Dim x, y, iLen As Integer
        Dim nombre As String = ""
        Dim delimitadores() As String = {"(R:"}   ' 
        Dim delimitadores1() As String = {"CN="}
        Dim delimitadores2() As String = {" "} ' 
        Dim vectoraux() As String
        Dim subvector() As String


        Dim valido As Boolean = False
        cadena = ""
        repre_name = ""
        repre_cif = ""

        vectoraux = subjet.Split(delimitadores, StringSplitOptions.None)
        If vectoraux.Count = 2 Then
            aux = vectoraux(0)
            subvector = aux.Split(delimitadores1, StringSplitOptions.None)
            If subvector.Count = 2 Then
                aux = subvector(1) ' ahora buscar primer blanco sera para el dni y resto nombre
                x = aux.IndexOf(" ")
                If x <> -1 Then
                    cadena = Trim(Mid(aux, 1, x + 1))
                    iLen = Len(cadena)
                    For y = 1 To iLen
                        If IsNumeric(Mid(cadena, y, 1)) Then
                            valido = True
                            Exit For
                        End If
                    Next y
                End If
                If valido Then
                    repre_cif = cadena
                    repre_name = Trim(Mid(aux, x + 1))
                Else
                    repre_cif = ""
                    repre_name = aux
                End If
            End If
        End If
    End Sub

    Public Function asigna_certificado_almacen(indice_certificado) As X509Certificate2

        Dim store As X509Store = New X509Store(StoreName.My, StoreLocation.CurrentUser)
        store.Open(OpenFlags.ReadOnly)
        Dim certificate As X509Certificate2 = store.Certificates.Item(indice_certificado)
        store.Close()

        Return certificate

    End Function

    Public Function carga_certificado(certificado_fich As String, passw As String) As X509Certificate2

        Dim certificate As New X509Certificate2(certificado_fich, passw)

        Return certificate

    End Function





End Module
