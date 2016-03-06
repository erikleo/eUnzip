Imports System.IO

Public Class eUnzip

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Dim zipFilePath As String
        Dim zipFileSourcedir As String
        Dim zipFileName As String

        zipFilePath = Command()

        'remove quotes on the end if they are there

        If Microsoft.VisualBasic.Left(zipFilePath, 1) = """" And Microsoft.VisualBasic.Right(zipFilePath, 1) = """" Then
            zipFilePath = Mid(zipFilePath, 2, Len(zipFilePath) - 2)
        End If


        If File.Exists(zipFilePath) Then
            zipFileSourcedir = System.IO.Path.GetDirectoryName(zipFilePath)
            zipFileName = System.IO.Path.GetFileNameWithoutExtension(zipFilePath)

            Dim sc As New Shell32.Shell()

            'Declare your input zip file as folder  .
            Dim input As Shell32.Folder = sc.NameSpace(zipFilePath)
            Dim output As Shell32.Folder

            'if there are more than one items in the root of the zip put it in a folder otherwise just unzip
            If input.Items.Count = 1 Then
                output = sc.NameSpace(zipFileSourcedir)
            Else
                System.IO.Directory.CreateDirectory(zipFileSourcedir & "\" & zipFileName)
                output = sc.NameSpace(zipFileSourcedir & "\" & zipFileName)
            End If

            'Extract the files from the zip file using the CopyHere command .
            output.CopyHere(input.Items, 0)


            'delete the original zip file
            File.Delete(zipFilePath)

            'note for some reason the zero lengh file would not delte from the desktop.. therefore call
            'a refresh
            Win32Helper.NotifyFileAssociationChanged()
        End If
        End
    End Sub
End Class