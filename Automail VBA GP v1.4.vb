Sub Send_emails()

Dim CDO_Mail As Object
Dim CDO_Config As Object 'íàõóÿ åãî ïèëèòü êàê îáäæåêò, åñëè ýòî õåðíÿ âñòðîåííà â áèáëèîòåêó? Îòâåò ïðîñò - ÿ äîëáàåá))))
Dim SMTP_Config As Variant
Dim strSubject As String
Dim strFrom As String
Dim strTo As String
Dim strCc As String
Dim strBcc As String
Dim strBody As String
Dim counter As Integer


For counter = 4 To 4 + Int(Sheet1.Cells(2, 4)) - 1
strSubject = Sheet1.Cells(counter, 5) + " " + Str(Sheet1.Cells(counter, 2)) + " " + Sheet1.Cells(counter, 3)
strFrom = Sheet1.Cells(2, 7)
strTo = Sheet1.Cells(2, 3)
strCc = ""
strBcc = ""
strBody = Sheet1.Cells(counter, 4)

Set CDO_Mail = CreateObject("CDO.Message")
With CDO_Mail
    .BodyPart.Charset = "utf-8" 'Êàê æå ÿ ÅÁÀË ýòîò VBA ñóóóóêààààà!!!! åÁÀÍÛå êîäèðîâêè 3 áëÿäñêèõ ÷àñà, ÷òîáû âûæàòü UTF-8 èç ýòîé õóéíè!!!! ÀÀÀÀÀÀÀ ñóêà
End With
On Error GoTo Error_Handling

Set CDO_Config = CreateObject("CDO.Configuration")
CDO_Config.Load -1

Set SMTP_Config = CDO_Config.Fields

With SMTP_Config
 .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
 .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
 .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
 .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Sheet1.Cells(2, 7)
 .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Sheet1.Cells(3, 7)
 .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = Sheet1.Cells(4, 7)
 .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
 .Update
 
 
End With


With CDO_Mail
 Set .Configuration = CDO_Config
End With

CDO_Mail.Subject = strSubject
CDO_Mail.From = strFrom
CDO_Mail.To = strTo
CDO_Mail.TextBody = strBody
CDO_Mail.CC = strCc
CDO_Mail.BCC = strBcc
CDO_Mail.Send

Error_Handling:
If Err.Description <> "" Then MsgBox Err.Description

Application.Wait (Now + TimeValue("0:00:05"))

Next counter

End Sub
