Imports System.IO
Public Module ExportThumbnailToDB

    Private oApprenticeApp As Inventor.ApprenticeServerComponent
    Private Doc As Inventor.ApprenticeServerDocument

    Function GetComponentPathsFromFile(File As Scripting.File) As Collection

        Dim stream As Scripting.TextStream
        stream = File.OpenAsTextStream(Scripting.IOMode.ForReading, Scripting.Tristate.TristateUseDefault)

        Dim pathsArray As New Collection

        While Not stream.AtEndOfStream
            pathsArray.Add(Split(stream.ReadLine, ";"))
        End While

        Return pathsArray

    End Function
    Function GetThumbnail(ComponentPath As String) As System.Drawing.Image

        Doc = oApprenticeApp.Open(ComponentPath)
        If Doc.Type <> 50340736 Then
            Throw New Exception("No se pudo abrir el documento con ruta: " & ComponentPath)
        End If

        Dim Thumb As stdole.IPictureDisp
        Thumb = Doc.Thumbnail

        Dim img As System.Drawing.Image
#Disable Warning BC40000 ' Type or member is obsolete
        img = Compatibility.VB6.IPictureDispToImage(Thumb)

        Return img

    End Function

    Function Main(ByVal cmdArgs() As String) As Integer

        If cmdArgs.Length < 3 Then
            Console.WriteLine("Faltan parametros: Ejecutar ./VB_scripts.exe [URL a lista de rutas de componentes] [Nombre de proyecto] [Ruta destino]")
            Return 1
        End If

        Dim fs As Scripting.FileSystemObject

        Try
            fs = CreateObject("Scripting.FileSystemObject")
        Catch ex As Exception
            Console.WriteLine("Error " & ex.Message & ": No se pudo crear un FileSystemObject, verificar la existencia de Microsoft Scripting Runtime.")
            Return 2
        End Try


        If System.IO.File.Exists(cmdArgs(0)) Then
            Console.WriteLine("No se encontró la lista de rutas de componentes")
        End If

        Dim componentPaths As Scripting.File
        componentPaths = fs.GetFile(cmdArgs(0))

        Console.WriteLine("Lista de rutas obtenida exitosamente")

        Dim destPath As String
        destPath = cmdArgs(2) & cmdArgs(1) & "\"

        If Not fs.FolderExists(destPath) Then
            Console.WriteLine("Carpeta de destino inexistente. Creandola...")
            fs.CreateFolder(destPath)
            Console.WriteLine("Carpeta de destino creada.")
        End If

        Dim componentsArray As Collection

        Console.WriteLine("Escaneando lista de rutas de componentes...")

        componentsArray = GetComponentPathsFromFile(componentPaths)

        Console.WriteLine("Lista de rutas de componentes escaneada.")

        Dim img As Drawing.Image

        If oApprenticeApp Is Nothing Then
            Console.WriteLine("Iniciando Apprentice Server...")
            oApprenticeApp = New Inventor.ApprenticeServerComponent
        End If

        If oApprenticeApp.Type <> 50341120 Then
            Console.WriteLine("No se pudo iniciar el Apprentice Server.")
            Return 2
        End If

        Console.WriteLine("Obteniendo Thumbnails de archivos...")

        For i = 1 To componentsArray.Count
            Try
                img = GetThumbnail(componentsArray(i)(1))
                Console.WriteLine("Imagen para " & componentsArray(i)(0) & " obtenida.")
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                Return 2
            End Try
            Try
                img.Save(destPath & componentsArray(i)(0) & ".png", Drawing.Imaging.ImageFormat.Png)
                Console.WriteLine("Guardada con nombre: " & componentsArray(i)(0) & ".png")
            Catch ex As Exception
                Console.WriteLine("No se pudo guardar el Thumbnail del componente " & componentsArray(i)(0) & ". No tiene un formato aceptable. Salteando...")
            End Try
        Next

        Console.WriteLine("Thumbnails exportados exitosamente a: " & Path.GetFullPath(destPath))

        Return 0

    End Function

End Module
