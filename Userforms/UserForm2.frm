VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Prompt"
   ClientHeight    =   9435.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10050
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DisplayText As String


Private Sub btnGenerate_Click()
         Dim request As Object
         Dim text As String, response As String, API As String, error_result As String
         Dim startPos As Long, endPos As Long, status_code As Long
         Dim EDisplayText As String
         Dim userInput As String
    
        userInput = rephrasedForm1.PromptBox.text
        
        'MsgBox "User Input: " & userInput, vbInformation
    
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
           rephrasedForm1.ChatGPTResponce.text = DisplayText
           rephrasedForm1.ChatGPTResponce.Value = DisplayText
        
           ' Check the user's response
       
         Else
              startPos = InStr(response, """message"": """) + Len("""message"": """)
              endPos = InStr(startPos, response, """")
              If startPos > Len("""message"": """) And endPos > startPos Then
                  DisplayText = Mid(response, startPos, endPos - startPos)

              Else
                  DisplayText = ""
              End If
        
          End If
        
    'Clean up the object
    Set request = Nothing

End Sub


Private Sub btnInsert2_Click()
    ' Replace the selected text with the response
     Dim selectedText As Range
     Set selectedText = Selection.Range
     selectedText.text = DisplayText
    rephrasedForm1.Hide
End Sub

Private Sub btnCancel2_Click()
    rephrasedForm1.Hide
End Sub



Private Sub btnCopy2_Click()
 
    ' Set the new text to the clipboard
    ClipBoard_SetData DisplayText
    

    ' Update the button caption
    btnCopy2.Caption = "Copied!"
    btnCopy2.BackColor = &H6B882C
    
    ' Pause for 2 seconds using a loop
    Dim endTime As Double
    endTime = Now + TimeValue("00:00:04")
    Do While Now < endTime
        DoEvents
    Loop
    btnCopy2.Caption = "Copy"
    btnCopy2.BackColor = &H97753F
    
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

