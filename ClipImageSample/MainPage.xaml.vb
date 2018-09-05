' The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

Imports Windows.Graphics.Imaging
Imports Windows.Storage
Imports Windows.Storage.Pickers
Imports Windows.Storage.Streams
''' <summary>
''' An empty page that can be used on its own or navigated to within a Frame.
''' </summary>
Public NotInheritable Class MainPage
    Inherits Page


    Public Shared Async Function SaveCroppedBitmapAsync(originalImageFile As StorageFile, newImageFile As StorageFile, startPoint As Point, cropSize As Size) As Task
        Dim startPointX As UInteger = CUInt(Math.Floor(startPoint.X))
        Dim startPointY As UInteger = CUInt(Math.Floor(startPoint.Y))
        Dim height As UInteger = CUInt(Math.Floor(cropSize.Height))
        Dim width As UInteger = CUInt(Math.Floor(cropSize.Width))
        Using originalImgFileStream As IRandomAccessStream = Await originalImageFile.OpenReadAsync()

            Dim decoder As BitmapDecoder = Await BitmapDecoder.CreateAsync(originalImgFileStream)

            If startPointX + width > decoder.PixelWidth Then
                startPointX = decoder.PixelWidth - width
            End If

            If startPointY + height > decoder.PixelHeight Then
                startPointY = decoder.PixelHeight - height
            End If

            Using newImgFileStream As IRandomAccessStream = Await newImageFile.OpenAsync(FileAccessMode.ReadWrite)
                Dim pixels As Byte() = Await GetPixelData(decoder, startPointX, startPointY, width, height, decoder.PixelWidth,
                    decoder.PixelHeight)

                Dim encoderID As New Guid
                encoderID = Guid.Empty

                Select Case newImageFile.FileType.ToLower()
                    Case ".png"
                        encoderID = BitmapEncoder.PngEncoderId
                        Exit Select
                    Case ".bmp"
                        encoderID = BitmapEncoder.BmpEncoderId
                        Exit Select
                    Case Else
                        encoderID = BitmapEncoder.JpegEncoderId
                        Exit Select
                End Select

                Dim propertySet As New BitmapPropertySet()

                If decoder.PixelWidth > 3000 Or decoder.PixelHeight > 3000 Then
                    Dim qualityValue As New BitmapTypedValue(0.4, PropertyType.Single)
                    propertySet.Add("ImageQuality", qualityValue)
                Else
                    Dim qualityValue As New BitmapTypedValue(0.7, PropertyType.Single)
                    propertySet.Add("ImageQuality", qualityValue)
                End If

                Dim bmpEncoder As BitmapEncoder = Await BitmapEncoder.CreateAsync(encoderID, newImgFileStream)

                Try
                    Await bmpEncoder.BitmapProperties.SetPropertiesAsync(propertySet)
                Catch ex As Exception
                    Debug.WriteLine(ex.Message)
                End Try


                ''''''''' Exception in this point,  pixel becomes null!!!! why????

                bmpEncoder.SetPixelData(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Straight, width, height, decoder.DpiX, decoder.DpiY,
                        pixels)

                Await bmpEncoder.FlushAsync()

            End Using
        End Using
    End Function


    Private Shared Async Function GetPixelData(decoder As BitmapDecoder, startPointX As UInteger, startPointY As UInteger, width As UInteger, height As UInteger, scaledWidth As UInteger,
    scaledHeight As UInteger) As Task(Of Byte())
        Dim transform As New BitmapTransform()
        Dim bounds As New BitmapBounds()
        bounds.X = startPointX
        bounds.Y = startPointY
        bounds.Height = height
        bounds.Width = width
        transform.Bounds = bounds
        transform.ScaledWidth = scaledWidth
        transform.ScaledHeight = scaledHeight
        Dim pix As PixelDataProvider = Await decoder.GetPixelDataAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Straight, transform, ExifOrientationMode.IgnoreExifOrientation, ColorManagementMode.ColorManageToSRgb)
        Dim pixels As Byte() = pix.DetachPixelData()
        Return pixels
    End Function

    Private Async Sub Button_Click(sender As Object, e As RoutedEventArgs)

        Dim picker As New FileOpenPicker()
        picker.FileTypeFilter.Add(".png")
        Dim source As StorageFile = Await picker.PickSingleFileAsync()

        Dim target As StorageFile = Await ApplicationData.Current.LocalFolder.CreateFileAsync("image.png")
        Dim point As New Point(10, 10)
        Dim size As New Size(40, 40)

        Await MainPage.SaveCroppedBitmapAsync(source, target, point, size)
    End Sub
End Class
