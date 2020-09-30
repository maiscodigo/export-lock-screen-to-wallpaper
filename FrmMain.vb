Imports System.IO

Public Class FrmMain
    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Exibe um modal para seleção de diretório
        Using fbdDest As New FolderBrowserDialog
            fbdDest.Description = "Selecione o diretório onde serão salvas as imagens:"
            fbdDest.RootFolder = Environment.SpecialFolder.UserProfile
            fbdDest.ShowNewFolderButton = True
            fbdDest.ShowDialog()
            SubGetFiles(fbdDest.SelectedPath)
        End Using

        'Finaliza a aplicação
        End
    End Sub

    Private Sub SubGetFiles(sDest As String)
        Dim di As DirectoryInfo = New DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets")

        'Percorre por todo o diretório em busca de arquivos
        For Each fi In di.GetFiles()
            'Converte o arquivo em uma matriz de bytes
            Dim bFileBytes() As Byte = My.Computer.FileSystem.ReadAllBytes(fi.FullName)

            'Pega os 4 primeiros bytes da matriz e os converte em hexa
            Dim sAssign As String = BitConverter.ToString(bFileBytes, 0, 4).Replace("-", "")

            'Verifica se a assinatura do arquivo atual é de um .jpg
            If sAssign.Contains("FFD8FFE0") Or sAssign.Contains("FFD8FFE1") Then
                Dim bmp As New Bitmap(fi.FullName)

                'Se a dimenção for igual ou maior que 1920 x 1080
                If bmp.Width >= 1920 And bmp.Height >= 1080 Then
                    'Copia para a pasta informada não sobreescrevendo caso o arquivo exista
                    My.Computer.FileSystem.CopyFile(fi.FullName, sDest & "\" & fi.Name & ".jpg", False)
                End If
            End If
        Next
    End Sub
End Class
