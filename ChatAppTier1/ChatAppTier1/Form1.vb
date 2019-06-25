Public Class Form1
    Dim clicked As Boolean = False

    Function GetUserName() As String
        If TypeOf My.User.CurrentPrincipal Is
      Security.Principal.WindowsPrincipal Then
            ' The application is using Windows authentication.
            ' The name format is DOMAIN\USERNAME.
            Dim parts() As String = Split(My.User.Name, "\")
            Dim username As String = parts(1)
            Return username
        Else
            ' The application is using custom authentication.
            Return My.User.Name
        End If

    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate("https://avaya.conversive.com/ABMI/chat/Chat.aspx")

        While WebBrowser1.ReadyState <> WebBrowserReadyState.Complete
            Application.DoEvents()
        End While

        WebBrowser1.Document.GetElementById("txtUserName").InnerText = GetUserName()
        WebBrowser1.Document.GetElementById("txtAccountName").InnerText = "Avaya"


        TimerSemChat.Start()


    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted

    End Sub

    Private Sub TimerDIV_Tick(sender As Object, e As EventArgs) Handles TimerDIV.Tick
        Try
            Dim contador As Integer
            contador = WebBrowser1.Document.GetElementsByTagName("h3").Count
            Label1.Text = contador.ToString()

        Catch ex As Exception
            Application.DoEvents()
        End Try

    End Sub

    Private Sub TimerSemChat_Tick(sender As Object, e As EventArgs) Handles TimerSemChat.Tick
        If Label1.Text = 13 Then
            Button1.PerformClick()
            TimerComChat.Start()
            TimerSemChat.Stop()
        Else
            Application.DoEvents()
        End If



    End Sub

    Private Sub TimerComChat_Tick(sender As Object, e As EventArgs) Handles TimerComChat.Tick
        If Label1.Text <> 13 Then
            TimerSemChat.Start()
            TimerComChat.Stop()
        Else
            Application.DoEvents()
        End If


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        clicked = True

        If clicked = True Then
            Try
                WebBrowser1.Document.GetElementById("txtInput").InnerText = ("Olá, Bem vindo ao Service Desk da Avaya Brasil! Como posso ajudar?")
                WebBrowser1.Document.GetElementById("btnSendInput").InvokeMember("Click")
                My.Computer.Audio.Play(My.Resources.notification, AudioPlayMode.Background)


                Me.WindowState = FormWindowState.Normal
                Me.Show()
                Me.TopMost = True



            Catch ex As Exception
                Beep()
                Application.DoEvents()
                MessageBox.Show("Mensagem Automatica não Enviada !! ",
         "ERRO", MessageBoxButtons.OK,
             MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)

            End Try
        End If

        clicked = False

    End Sub



    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        T2Tool.Show()
        T2Tool.BringToFront()
    End Sub
End Class
