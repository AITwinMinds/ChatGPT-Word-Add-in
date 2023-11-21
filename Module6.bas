Attribute VB_Name = "Module6"
Option Explicit


Public Sub thePrompt(control As IRibbonControl)

    Dim request As Object
    Dim text As String, response As String, API As String, DisplayText As String, error_result As String
    Dim startPos As Long, endPos As Long, status_code As Long
    Dim selectedText As Range
    Dim EDisplayText As String
    Dim userInput As String
    
    userInput = InputBox("Enter your Prompt:", "ASK ChatGPT")
    
    If userInput <> "" Then
        text = userInput
    Else
        MsgBox "Prompt cannot be blank!", vbExclamation
    End If
    
    'API Info
    API = "https://api.openai.com/v1/chat/completions"

    If api_key = "" Then
        MsgBox "Error: API key is blank!"
        Exit Sub
    End If
    

    
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
      DisplayText = ExtractContent(response)
        
     'Insert response text into Word document and format it in red color
     ' Display the result in the Word document
    Selection.TypeParagraph ' Add a new paragraph before inserting the text
    Selection.TypeText "Your prompt: " & vbCrLf & text & vbCrLf
    Selection.TypeParagraph ' Add a new paragraph before inserting the text
    Selection.TypeText "ChatGPT: " & vbCrLf & DisplayText
    
    Else
        startPos = InStr(response, """message"": """) + Len("""message"": """)
        endPos = InStr(startPos, response, """")
        If startPos > Len("""message"": """) And endPos > startPos Then
            DisplayText = Mid(response, startPos, endPos - startPos)

        Else
            DisplayText = ""
        End If
        
        'Insert error message into Word document
        EDisplayText = "Error : " & DisplayText
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



