Imports System.Drawing.Imaging

Public Class Form1
    Private Sub VaryQualityLevel()
        ' Get a bitmap. 
        Dim bmp1 As New Bitmap("c:\TestPhoto.jpg")
        Dim jgpEncoder As ImageCodecInfo = GetEncoder(ImageFormat.Jpeg)

        ' Create an Encoder object based on the GUID 
        ' for the Quality parameter category. 
        Dim myEncoder As System.Drawing.Imaging.Encoder = System.Drawing.Imaging.Encoder.Quality

        ' Create an EncoderParameters object. 
        ' An EncoderParameters object has an array of EncoderParameter 
        ' objects. In this case, there is only one 
        ' EncoderParameter object in the array. 
        Dim myEncoderParameters As New EncoderParameters(1)

        Dim myEncoderParameter As New EncoderParameter(myEncoder, 20)
        myEncoderParameters.Param(0) = myEncoderParameter
        bmp1.Save("c:\TestPhotoQuality20.jpg", jgpEncoder, myEncoderParameters)

    End Sub

    Private Function GetEncoder(ByVal format As ImageFormat) As ImageCodecInfo

        Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageDecoders()

        Dim codec As ImageCodecInfo
        For Each codec In codecs
            If codec.FormatID = format.Guid Then
                Return codec
            End If
        Next codec
        Return Nothing

    End Function

    
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        VaryQualityLevel()
        Me.Close()
    End Sub
End Class
