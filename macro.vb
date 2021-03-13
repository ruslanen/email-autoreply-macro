' Адрес электронной почты, с которой необходимо производить автоответ
Private Const MailAddress As String = ""

Private WithEvents Items As Outlook.Items
Private Sub Application_Startup()
  Dim olApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Set olApp = Outlook.Application
  Set objNS = olApp.GetNamespace("MAPI")
  Set Items = objNS.GetDefaultFolder(olFolderInbox).Items
End Sub

' Обработчик события получения нового письма.
' Проверяет соответствие адреса электронной почты указанному и вызывает обработку.
Private Sub Items_ItemAdd(ByVal item As Object)
  On Error GoTo ErrorHandler
  Dim Msg As Outlook.MailItem
  If TypeName(item) = "MailItem" Then
    If Not (item.Sender Is Nothing) Then
        If (item.SenderEmailAddress = MailAddress) Then
            ' If Debug
            ' MsgBox item.Body
            Call ProcessMessage(item.Body)
        End If
    End If
  End If
ProgramExit:
  Exit Sub
ErrorHandler:
  MsgBox Err.Number & " - " & Err.Description
  Resume ProgramExit
End Sub

' Обрабатывает сообщение и производит автоответ.
Public Sub ProcessMessage(Text As String)
    Dim MailFromPosition, MailToPosition
    Dim SubjectFromPosition, SubjectToPosition
    Dim Mail, Subject
    ' GetMail
    ' 7 = mailto:
    MailFromPosition = InStr(1, Text, "mailto:") + 7
    MailToPosition = InStr(MailFromPosition, Text, "?")
    Mail = Mid(Text, MailFromPosition, MailToPosition - MailFromPosition)
    ' GetSubject
    ' 8 = subject=
    SubjectFromPosition = InStr(MailToPosition, Text, "subject=") + 8
    ' Понимается, что в письме будет гиперссылка вида:
    ' <a href = mailto:example@example.com?subject=test_12345678_123456
    ' Body=Несоглаcие>Нет</a>
    ' (сохранен разрыв строки)
    SubjectToPosition = InStr(SubjectFromPosition, Text, "Body")
    Subject = Mid(Text, SubjectFromPosition, SubjectToPosition - SubjectFromPosition)

    Call SendEmail(Mail, Subject)
End Sub

' Отправляет письмо указанному получателю по адресу электронной почты с указанной темой
Public Sub SendEmail(MailTo, Subject)
    Dim Email As Outlook.Application
    Set Email = New Outlook.Application
    Dim newMail As Outlook.MailItem
    Set newMail = Email.CreateItem(olMailItem)
    newMail.To = MailTo
    newMail.Subject = Subject
    newMail.Send
End Sub

