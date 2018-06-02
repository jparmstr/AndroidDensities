Imports System.IO
Imports System.IO.Compression
Imports System.Drawing

Module Module1

    ' This application automatically creates pixel density variations for Android apps
    ' Source images should be in \res\drawable
    ' Resized images will be saved in their corresponding pixel density drawable directories
    ' Source images will be moved to ARCHIVE_DIRECTORY_NAME then zipped so that Android Studio doesn't complain about them

    ' .exe should be located in the \res\ directory

    ' Constants
    ' Screen density multipliers assuming that xxxhdpi files are provided (see android pixel densities.xlsx)
    Private ReadOnly XXXHDPI As Double = 1.0
    Private ReadOnly XXHDPI As Double = 0.75
    Private ReadOnly XHDPI As Double = 0.5
    Private ReadOnly HDPI As Double = 0.375
    Private ReadOnly MDPI As Double = 0.25
    Private ReadOnly LDPI As Double = 0.1875

    ' Drawable resource directory names
    Private ReadOnly SOURCE_DIRECTORY_NAME As String = "\drawable"
    Private ReadOnly ARCHIVE_DIRECTORY_NAME As String = "\original assets"
    Private ReadOnly XXXHDPI_NAME As String = "\drawable-xxxhdpi"
    Private ReadOnly XXHDPI_NAME As String = "\drawable-xxhdpi"
    Private ReadOnly XHDPI_NAME As String = "\drawable-xhdpi"
    Private ReadOnly HDPI_NAME As String = "\drawable-hdpi"
    Private ReadOnly MDPI_NAME As String = "\drawable-mdpi"
    Private ReadOnly LDPI_NAME As String = "\drawable-ldpi"

    ' Settings
    Private setting_interpolation_mode = Drawing2D.InterpolationMode.HighQualityBicubic
    ' note: file extensions for image formats are defined in saveFile()
    Private setting_image_format = Imaging.ImageFormat.Png

    Sub Main()
        ' Get directory of executable
        Dim basePath = My.Application.Info.DirectoryPath

        ' Detect the .exe being in the wrong place
        If Not Directory.Exists(basePath + SOURCE_DIRECTORY_NAME) Then
            Console.WriteLine("Could not find the source directory (" + SOURCE_DIRECTORY_NAME +
                              "). Make sure the .exe is located in the \res\ folder and your source " +
                              "images are located in \res" + SOURCE_DIRECTORY_NAME)
            Console.ReadLine()
            Exit Sub
        End If

        ' Get enumerable of .jpg or .png images in current directory
        Dim imageFiles = From file In Directory.EnumerateFiles(basePath + SOURCE_DIRECTORY_NAME)
                         Where file.ToLower().EndsWith(".jpg") Or
                             file.ToLower.EndsWith(".png")

        ' Detect "no files found" error
        If imageFiles.Count = 0 Then
            Console.WriteLine("Found no images in " + SOURCE_DIRECTORY_NAME + ". " +
                              "Make sure the .exe is located in the \res\ folder and your source " +
                              "images are located in \res" + SOURCE_DIRECTORY_NAME + "." + vbNewLine +
                              "Supported input formats are *.jpg and *.png")
            Console.ReadLine()
            Exit Sub
        End If

        ' Create a list of all pixel density multipliers
        Dim pixelDensityMultipliers As New List(Of Double)
        With pixelDensityMultipliers
            .Add(XXXHDPI)
            .Add(XXHDPI)
            .Add(XHDPI)
            .Add(HDPI)
            .Add(MDPI)
            .Add(LDPI)
        End With

        ' Create directories if they don't exist
        If Not Directory.Exists(basePath + ARCHIVE_DIRECTORY_NAME) Then Directory.CreateDirectory(basePath + ARCHIVE_DIRECTORY_NAME)
        If Not Directory.Exists(basePath + XXXHDPI_NAME) Then Directory.CreateDirectory(basePath + XXXHDPI_NAME)
        If Not Directory.Exists(basePath + XXHDPI_NAME) Then Directory.CreateDirectory(basePath + XXHDPI_NAME)
        If Not Directory.Exists(basePath + XHDPI_NAME) Then Directory.CreateDirectory(basePath + XHDPI_NAME)
        If Not Directory.Exists(basePath + HDPI_NAME) Then Directory.CreateDirectory(basePath + HDPI_NAME)
        If Not Directory.Exists(basePath + MDPI_NAME) Then Directory.CreateDirectory(basePath + MDPI_NAME)
        If Not Directory.Exists(basePath + LDPI_NAME) Then Directory.CreateDirectory(basePath + LDPI_NAME)

        Dim counter As Integer = 0
        Dim counterTotal As Integer = imageFiles.Count

        ' Process all images from basePath directory
        For Each i In imageFiles
            ' Store the file name
            Dim fileName = getFileName_fromFullPath(i)

            ' Open file as Bitmap image
            Dim thisImageBitmap = getBitmapFromFile(i)

            ' Create a list to hold resized bitmaps
            Dim resizedBitmaps As New List(Of Bitmap)

            ' Create pixel density variations
            For Each pdMultiplier In pixelDensityMultipliers
                resizedBitmaps.Add(getResizedBitmap(thisImageBitmap, pdMultiplier))
            Next

            ' Save the files
            saveFile(resizedBitmaps.Item(0), basePath, XXXHDPI_NAME, fileName)
            saveFile(resizedBitmaps.Item(1), basePath, XXHDPI_NAME, fileName)
            saveFile(resizedBitmaps.Item(2), basePath, XHDPI_NAME, fileName)
            saveFile(resizedBitmaps.Item(3), basePath, HDPI_NAME, fileName)
            saveFile(resizedBitmaps.Item(4), basePath, MDPI_NAME, fileName)
            saveFile(resizedBitmaps.Item(5), basePath, LDPI_NAME, fileName)

            ' Move the original file to ARCHIVE_DIRECTORY_NAME
            Dim sourceImageDestination = basePath + ARCHIVE_DIRECTORY_NAME + "\" + fileName
            Try
                If File.Exists(sourceImageDestination) Then File.Delete(sourceImageDestination)
                File.Move(i, sourceImageDestination)
            Catch e As Exception
                Console.WriteLine(e.Message)
                Console.ReadLine()
            End Try

            ' Print progress to console application window
            counter += 1
            Console.Clear()
            Console.WriteLine(counter.ToString + " of " + counterTotal.ToString)
        Next

        ' Compress the original assets directory so that Android Studio doesn't complain
        Dim zipName = basePath + ARCHIVE_DIRECTORY_NAME + ".zip"

        ' Handle cases where there's a previous originalAssets.zip file (don't overwrite it)
        Dim safetyCounter As Integer = 0
        While File.Exists(zipName)
            safetyCounter += 1
            zipName = basePath + ARCHIVE_DIRECTORY_NAME + "-" + safetyCounter.ToString + ".zip"
        End While

        ' Create the zip file
        ZipFile.CreateFromDirectory(basePath + ARCHIVE_DIRECTORY_NAME, zipName)

        ' Delete the original assets folder (it has now been archived)
        Try
            Directory.Delete(basePath + ARCHIVE_DIRECTORY_NAME, True)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Sub saveFile(resizedBitmap As Bitmap, basePath As String, densityName As String, fileName As String)
        ' Choose the file extension based on setting_image_format
        Dim fileExtension As String = ""
        Select Case setting_image_format.ToString
            Case Imaging.ImageFormat.Png.ToString
                fileExtension = ".png"
            Case Imaging.ImageFormat.Jpeg.ToString
                fileExtension = ".jpg"
            Case Else
                fileExtension = ".error"
        End Select

        ' Build the save path
        Dim savePath = basePath + densityName + "\" + getFileNameWithoutExtension(fileName) + fileExtension

        ' Save the file
        resizedBitmap.Save(savePath, setting_image_format)
    End Sub



    Function getBitmapFromFile(fullPath As String) As Bitmap
        Return New Bitmap(Bitmap.FromStream(New MemoryStream(File.ReadAllBytes((fullPath)))))
    End Function

    Function getResizedBitmap(original As Bitmap, multiplier As Double) As Bitmap
        ' Intentionally discarding remainders
        ' because I've had Android Studio complain about images whose sizes were rounded normally
        ' and Android Studio seemed to like an online tool which always rounded down

        Dim newWidth As Integer = Math.Floor(original.Width * multiplier)
        Dim newHeight As Integer = Math.Floor(original.Height * multiplier)

        Dim resizedBitmap As New Bitmap(newWidth, newHeight)

        ' Draw the original image onto the resized image
        Dim g As Graphics = Graphics.FromImage(resizedBitmap)
        g.InterpolationMode = setting_interpolation_mode
        g.DrawImage(original, New Rectangle(0, 0, resizedBitmap.Width, resizedBitmap.Height), New Rectangle(0, 0, original.Width, original.Height), GraphicsUnit.Pixel)
        g.Dispose()

        Return resizedBitmap
    End Function

    Function getFileName_fromFullPath(fullPath As String) As String
        Dim result As String
        Dim snipStart As Integer
        Dim snipEnd As Integer
        Dim snipLength As Integer

        snipStart = InStrRev(fullPath, "\") + 1
        snipEnd = Len(fullPath)
        snipLength = snipEnd - snipStart + 1

        result = Mid(fullPath, snipStart, snipLength)

        Return result
    End Function

    Function getDirectory(ByVal filepath As String) As String
        getDirectory = filepath.Substring(0, filepath.LastIndexOf("\") + 1)
    End Function

    Function getFileNameWithoutExtension(ByVal fileName As String) As String
        If fileName.Contains("\") Then
            fileName = getFileName_fromFullPath(fileName)
        End If

        Dim snipStart As Integer
        Dim snipEnd As Integer
        Dim snipLength As Integer

        snipEnd = InStrRev(fileName, ".") - 1
        snipStart = 1
        snipLength = Len(fileName) - (Len(fileName) - snipEnd)

        Return Mid(fileName, snipStart, snipLength)
    End Function

End Module
