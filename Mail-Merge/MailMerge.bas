Attribute VB_Name = "Main"
Option Explicit

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub MailMerge()
    Application.ScreenUpdating = False
    
    Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("MailMerge")
    Dim r As Long: r = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim outlookApp As Object: Set outlookApp = CreateObject("Outlook.Application")
    Dim outlookMail As Object
    
    Dim body_text As String
    body_text = "Hello," & vbNewLine & vbNewLine & _
                "This is a lengthy corporate email that was once sent to hundreds of people manually." & vbNewLine & vbNewLine & _
                "Sincerely," & vbNewLine & _
                "Bob"
    
    Dim i As Long
    For i = 2 To r
        '0 for olMailItem
        Set outlookMail = outlookApp.CreateItem(0)
        With outlookMail
            .To = WS.Range("A" & i).Value2 & ";" & WS.Range("B" & i).Value2
            .CC = WS.Range("C" & i).Value2 & ";" & WS.Range("D" & i).Value2
            .Subject = "Automation"
            '1 for olFormatPlain
            .BodyFormat = 1
            .Body = body_text
            'Could also add an attachment
            '.Attachments.Add ThisWorkbook.Path & "\" & "Important Annoucement.pdf"
            '.Save to save to Drafts folder, .Send to send immediately
            .Save
        End With
        Sleep (100)
    Next i
    
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    Application.ScreenUpdating = True
    MsgBox "Done!"
End Sub
