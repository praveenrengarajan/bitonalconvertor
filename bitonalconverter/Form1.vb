Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices

Public Class Form1
    Dim optimgname As Bitmap
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
        Dim filname As String = OpenFileDialog1.FileName
        Dim fname As Bitmap = New Bitmap(filname)
        PictureBox1.Image = fname
        optimgname = ConvertToBitonal(fname)
        PictureBox2.Image = optimgname
    End Sub


    Public Function ConvertToBitonal(ByRef original As Bitmap) As Bitmap
        Dim source As Bitmap = Nothing

        ' If original bitmap is not already in 32 BPP, ARGB format, then convert
        If original.PixelFormat <> PixelFormat.Format32bppArgb Then
            source = New Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb)
            source.SetResolution(original.HorizontalResolution, original.VerticalResolution)
            Using g As Graphics = Graphics.FromImage(source)
                g.DrawImageUnscaled(original, 0, 0)
            End Using
        Else
            source = original
        End If

        ' Lock source bitmap in memory
        Dim sourceData As BitmapData = source.LockBits(New Rectangle(0, 0, source.Width, source.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb)

        ' Copy image data to binary array
        Dim imageSize As Integer = sourceData.Stride * sourceData.Height
        Dim sourceBuffer(imageSize - 1) As Byte
        Marshal.Copy(sourceData.Scan0, sourceBuffer, 0, imageSize)

        ' Unlock source bitmap
        source.UnlockBits(sourceData)

        ' Create destination bitmap
        Dim destination As Bitmap = New Bitmap(source.Width, source.Height, PixelFormat.Format1bppIndexed)
        destination.SetResolution(original.HorizontalResolution, original.VerticalResolution)

        ' Lock destination bitmap in memory
        Dim destinationData As BitmapData = destination.LockBits(New Rectangle(0, 0, destination.Width, destination.Height), ImageLockMode.WriteOnly, PixelFormat.Format1bppIndexed)

        ' Create destination buffer
        imageSize = destinationData.Stride * destinationData.Height

        Dim destinationBuffer(imageSize - 1) As Byte
        Dim sourceIndex As Integer = 0
        Dim destinationIndex As Integer = 0
        Dim pixelTotal As Integer = 0
        Dim destinationValue As Byte = 0
        Dim pixelValue As Integer = 128
        Dim height As Integer = source.Height
        Dim width As Integer = source.Width
        Dim threshold As Integer = 500

        ' Iterate lines
        For y As Integer = 0 To height - 1

            sourceIndex = y * sourceData.Stride
            destinationIndex = y * destinationData.Stride
            destinationValue = 0
            pixelValue = 128

            ' Iterate pixels
            For x As Integer = 0 To width - 1
                ' Compute pixel brightness (i.e. total of Red, Green, and Blue values) - Thanks murx
                ' B G R
                pixelTotal = CInt(sourceBuffer(sourceIndex)) + CInt(sourceBuffer(sourceIndex + 1)) + CInt(sourceBuffer(sourceIndex + 2))
                If (pixelTotal > threshold) Then
                    destinationValue += CType(pixelValue, Byte)
                End If
                If pixelValue = 1 Then
                    destinationBuffer(destinationIndex) = destinationValue
                    destinationIndex += 1
                    destinationValue = 0
                    pixelValue = 128
                Else
                    pixelValue >>= 1
                End If
                sourceIndex += 4
            Next
            If pixelValue <> 128 Then
                destinationBuffer(destinationIndex) = destinationValue
            End If
        Next

        ' Copy binary image data to destination bitmap
        Marshal.Copy(destinationBuffer, 0, destinationData.Scan0, imageSize)

        ' Unlock destination bitmap
        destination.UnlockBits(destinationData)

        ' Dispose of source if not originally supplied bitmap
        If Not Object.ReferenceEquals(source, original) Then
            source.Dispose()
        End If

        Return destination
    End Function

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        SaveFileDialog1.ShowDialog()
        Dim optfile As String = SaveFileDialog1.FileName
        Dim frmt As ImageFormat = ImageFormat.Tiff
        optimgname.Save(optfile, frmt)
    End Sub
End Class
