Attribute VB_Name = "Module2"
Option Explicit

Public Sub theReply(control As IRibbonControl)

    Dim request As Object
    Dim text As String, response As String, API As String, DisplayText7 As String, error_result As String
    Dim startPos As Long, endPos As Long, status_code As Long
    Dim prompt As String
    Dim selectedText As Range
    Dim EDisplayText As String

    'API Info
    API = "https://api.openai.com/v1/chat/completions"
    
    If api_key = "" Then
        MsgBox "Error: API key is blank!"
        Exit Sub
    End If
    
    ' Prompt the user to select text in the document
    If Selection.Type <> wdSelectionIP Then
        prompt = Trim(Selection.text)
        Set selectedText = Selection.Range
    Else
        MsgBox "Please select some text before running this macro."
        Exit Sub
    End If
    
    ' Add your additional prompt before the selected text
    Dim additionalPrompt As String
    additionalPrompt = "Write a reply to the following email. Avoid wordy language: "
    
    
        
    'Cleaning
    text = Replace(prompt, Chr(34), Chr(39))
       ' Combine the additional prompt and the selected text
    text = additionalPrompt & vbCrLf & vbCrLf & text
    text = Replace(text, vbLf, "")
    text = Replace(text, vbCr, "")
    text = Replace(text, vbCrLf, "")
 

    ' Remove selection
    Selection.Collapse

    'Create an HTTP request object
    Set request = CreateObject("MSXML2.XMLHTTP")
    With request
        .Open "POST", API, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & api_key
        .send "{""model"": ""gpt-3.5-turbo"",  ""messages"": [{""content"":""" & text & """,""role"":""user""}]," _
             & """temperature"": 1}"
      status_code = .Status
      response = .responseText
    End With

    'Extract content
    If status_code = 200 Then
     DisplayText7 = ExtractContent(response)

     DisplayText7 = ExtractContent(response)
      Dim rephrasedForm7 As New UserForm1
      rephrasedForm7.lblRephrasedText.text = DisplayText7
      rephrasedForm7.lblRephrasedText.Value = DisplayText7
      rephrasedForm7.Show vbModal
      
' Check the user's response
    Select Case rephrasedForm7.Tag
        Case "Insert"
            ' Replace the selected text with the response
            selectedText.text = DisplayText7
        Case "Cancel"
            Exit Sub
    End Select


        
    Else
        startPos = InStr(response, """message"": """) + Len("""message"": """)
        endPos = InStr(startPos, response, """")
        If startPos > Len("""message"": """) And endPos > startPos Then
            DisplayText7 = Mid(response, startPos, endPos - startPos)

        Else
            DisplayText7 = ""
        End If
        
        'Insert error message into Word document
        EDisplayText = "Error : " & DisplayText7
        selectedText.InsertAfter vbNewLine & EDisplayText
        
    End If
    
    
    'Clean up the object
    Set request = Nothing

End Sub



Function ExtractContent(jsonString As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim Content As String
    
    startPos = InStr(jsonString, """content"": """) + Len("""content"": """)
    endPos = InStr(startPos, jsonString, "},") - 2
    Content = Mid(jsonString, startPos, endPos - startPos)
    Content = Trim(Replace(Content, "\""", Chr(34)))
        
    Content = Replace(Content, vbCrLf, "")
    Content = Replace(Content, vbLf, "")
    Content = Replace(Content, vbCr, "")
    Content = Replace(Content, "\n", vbCrLf)
     
    If Right(Content, 1) = """" Then
      Content = Left(Content, Len(Content) - 1)
    End If
    
    ExtractContent = Content

End Function

