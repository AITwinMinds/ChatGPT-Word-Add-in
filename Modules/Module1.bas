Attribute VB_Name = "Module1"
Option Explicit
Public api_key As String
Public doc As Document
Public SelectedRephraseStyleIndex As Integer
Public RephraseStyle As String

' Callback to set the API key
Public Sub SetApiKey(control As IRibbonControl)
    
       ' Declare a document variable
    Dim currentDoc As Document
    ' Set the document variable to the active document
    Set currentDoc = ActiveDocument
    
    ' If the API key is blank, prompt the user to enter it
    api_key = InputBox("Enter your API key:", "Set API Key")

    ' Create the custom document property if it doesn't exist
    If Not CustomDocumentPropertyExists(currentDoc, "APIKey") Then
        currentDoc.CustomDocumentProperties.Add Name:="APIKey", LinkToContent:=False, _
            Type:=msoPropertyTypeString, Value:=api_key
    End If

    ' Save the entered API key to the document property
    If api_key <> "" Then
        currentDoc.CustomDocumentProperties("APIKey").Value = api_key
        MsgBox "API key set successfully!", vbInformation
    Else
        MsgBox "API key cannot be blank!", vbExclamation
    End If


End Sub

' Function to check if a custom document property exists
Function CustomDocumentPropertyExists(currentDoc As Document, propertyName As String) As Boolean
    Dim prop As DocumentProperty
    CustomDocumentPropertyExists = False

    For Each prop In currentDoc.CustomDocumentProperties
        If prop.Name = propertyName Then
            CustomDocumentPropertyExists = True
            Exit Function
        End If
    Next prop
End Function


' Callback function for setting default index for ToLanguageDropdown
Sub GetSelectedRephraseStyleIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = SelectedRephraseStyleIndex
End Sub

Sub DropDown_OnAction_rephraseStyle(control As IRibbonControl, id As String, index As Integer)
    ' Handle dropdown selection changes
    RephraseStyle = id
           ' Declare a document variable
    Dim currentDoc2 As Document
    ' Set the document variable to the active document
    Set currentDoc2 = ActiveDocument
    ' Create the custom document property if it doesn't exist
    If Not CustomDocumentPropertyExists(currentDoc2, "RephraseStyle") Then
        currentDoc2.CustomDocumentProperties.Add Name:="RephraseStyle", LinkToContent:=False, _
            Type:=msoPropertyTypeString, Value:=RephraseStyle
    End If
    ' Save the entered fromLanguage to the document property
    If RephraseStyle <> "" Then
        currentDoc2.CustomDocumentProperties("RephraseStyle").Value = RephraseStyle
    End If
 

End Sub



Public Sub chatGPT(control As IRibbonControl)

    Dim request As Object
    Dim text As String, response As String, API As String, DisplayText As String, error_result As String
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
    
    Select Case RephraseStyle
     Case "Simplify"
         additionalPrompt = "Paraphrase the paragraph below. Simplify the language and maintain the core ideas: "
     Case "Informal"
         additionalPrompt = "Rewrite the paragraph below in a more informal tone, without changing the core message: "
     Case "Professional"
         additionalPrompt = "Can you suggest a different way to phrase paragraph below to make it sound more professional?: "
     Case "Formal"
         additionalPrompt = "Rewrite the paragraph below in a more formal tone, without changing the core message: "
     Case "Generalize"
         additionalPrompt = "Rewrite the technical paragraph below in simpler language, making it easily understandable for a general audience: "
    End Select
    
  
        
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
      DisplayText = ExtractContent(response)
      
          ' Create and show the UserForm
      Dim rephrasedForm As New UserForm1
      rephrasedForm.lblRephrasedText.text = DisplayText
      rephrasedForm.lblRephrasedText.Value = DisplayText
      rephrasedForm.Show vbModal
      
' Check the user's response
    Select Case rephrasedForm.Tag
        Case "Insert"
            ' Replace the selected text with the response
            selectedText.text = DisplayText
        Case "Cancel"
            Exit Sub
    End Select
      
      ' Insert response text into Word document and format it in red color
      'Dim responseRange As Range
      'Set responseRange = selectedText.Duplicate
      'responseRange.Collapse wdCollapseEnd
      'responseRange.InsertAfter vbNewLine & DisplayText
      'responseRange.Font.Color = RGB(255, 0, 0) ' Red color

        
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

