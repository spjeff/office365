' from https://stackoverflow.com/questions/38200239/outlook-macro-reply-to-sender-with-template
Sub JobApply()

Dim origEmail As MailItem
Dim replyEmail As MailItem
Dim firstName As String

Set origEmail = ActiveExplorer.Selection(1)
Set replyEmail = CreateItemFromTemplate("C:\BIN\template.oft")
firstName = Split(origEmail.Reply.To, " ")(0)

replyEmail.To = origEmail.Reply.To & "<" & origEmail.SenderEmailAddress & ">"

replyEmail.HTMLBody = Replace(replyEmail.HTMLBody, "{0}", firstName) & origEmail.Reply.HTMLBody
'replyEmail.SentOnBehalfOfName = "email@domain.com"
replyEmail.Subject = "RE: " & origEmail.Subject
replyEmail.Recipients.ResolveAll
replyEmail.Display

Set origEmail = Nothing
Set replyEmail = Nothing

End Sub
