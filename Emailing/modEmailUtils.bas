Attribute VB_Name = "modEmailUtils"
Option Explicit


Public Type TemailInfo
    Sender As String
    DateReceived As Date
    Subject As String
    AttachementCount As Integer
End Type


Public Function getEmailInfo(omail As Outlook.MailItem) As TemailInfo
    Dim emailInfo As TemailInfo
    With omail
        getEmailInfo.Sender = .SenderName
        getEmailInfo.DateReceived = .ReceivedTime
        getEmailInfo.Subject = .Subject
        getEmailInfo.AttachementCount = .Attachments.Count
    End With
End Function
