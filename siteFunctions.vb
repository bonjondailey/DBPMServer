Imports System.Net
Imports System.Net.Mail


Public Class siteFunctions

    Public Function SendMail() As Boolean
        Dim retVal As Boolean = False
        Dim smtpClient As New SmtpClient("smtp.myserver.com", 587) 'I tried using different hosts and ports
        smtpClient.UseDefaultCredentials = False
        smtpClient.Credentials = New NetworkCredential("username@domain.com", "password")
        smtpClient.EnableSsl = True 'Also tried setting this to false

        Dim mm As New MailMessage
        mm.From = New MailAddress("username@domain.com")
        mm.Subject = "Test Mail"
        mm.IsBodyHtml = True
        mm.Body = "<h1>This is a test email</h1>"
        mm.To.Add("someone@domain.com")

        Try
            smtpClient.Send(mm)
            retVal = True
        Catch ex As Exception

        End Try

        mm.Dispose()
        smtpClient.Dispose()

        Return retVal
    End Function
End Class
