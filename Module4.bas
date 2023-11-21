Attribute VB_Name = "Module4"
Option Explicit

Public fromLanguage As String
Public toLanguage As String

Sub Ribbon_Load(ribbon As IRibbonUI)
    ' Initialize the Ribbon UI
End Sub

Sub DropDown_OnAction_from(control As IRibbonControl, id As String, index As Integer)
    ' Handle dropdown selection changes
    fromLanguage = id
    ' Print the selected option immediately
End Sub

Sub DropDown_OnAction_to(control As IRibbonControl, id As String, index As Integer)
    ' Handle dropdown selection changes
    toLanguage = GetLanguageName(id)
    
    ' Print the selected option immediately
End Sub

Private Sub Document_Open()
    ' Initialize the Ribbon UI
    Ribbon_Load ThisDocument.Application.CommandBars.GetRibbonUI
End Sub


Public Sub theTranslate(control As IRibbonControl)
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
    
    ' Check if fromLanguage and toLanguage are selected
    If fromLanguage = "" Or toLanguage = "" Then
        MsgBox "Error: Please select both 'From Language' and 'To Language' before translating."
        Exit Sub
    End If
    
    additionalPrompt = "Translate from " & fromLanguage & " to " & toLanguage & vbCrLf

        
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

     'Insert response text into Word document and format it in red color
      Dim responseRange As Range
      Set responseRange = selectedText.Duplicate
      responseRange.Collapse wdCollapseEnd
      responseRange.InsertAfter vbNewLine & DisplayText
      responseRange.Font.Color = RGB(255, 0, 0) ' Red color


        
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


Function GetLanguageName(languageId As String) As String
    ' Map language IDs to language names
    Select Case languageId
        Case "AfrikaansLanguage"
            GetLanguageName = "Afrikaans"
        Case "AlbanianLanguage"
            GetLanguageName = "Albanian"
        Case "ArabicLanguage"
            GetLanguageName = "Arabic"
        Case "ArmenianLanguage"
            GetLanguageName = "Armenian"
        Case "AzerbaijaniLanguage"
            GetLanguageName = "Azerbaijani"
        Case "BengaliLanguage"
            GetLanguageName = "Bengali"
        Case "BulgarianLanguage"
            GetLanguageName = "Bulgarian"
        Case "ChineseLanguage"
            GetLanguageName = "Chinese"
        Case "CroatianLanguage"
            GetLanguageName = "Croatian"
        Case "CzechLanguage"
            GetLanguageName = "Czech"
        Case "DanishLanguage"
            GetLanguageName = "Danish"
        Case "DutchLanguage"
            GetLanguageName = "Dutch"
        Case "EnglishLanguage"
            GetLanguageName = "English"
        Case "EstonianLanguage"
            GetLanguageName = "Estonian"
        Case "FijianLanguage"
            GetLanguageName = "Fijian"
        Case "FinnishLanguage"
            GetLanguageName = "Finnish"
        Case "FrenchLanguage"
            GetLanguageName = "French"
        Case "GeorgianLanguage"
            GetLanguageName = "Georgian"
        Case "GermanLanguage"
            GetLanguageName = "German"
        Case "GreekLanguage"
            GetLanguageName = "Greek"
        Case "HebrewLanguage"
            GetLanguageName = "Hebrew"
        Case "HindiLanguage"
            GetLanguageName = "Hindi"
        Case "HungarianLanguage"
            GetLanguageName = "Hungarian"
        Case "IcelandicLanguage"
            GetLanguageName = "Icelandic"
        Case "IndonesianLanguage"
            GetLanguageName = "Indonesian"
        Case "ItalianLanguage"
            GetLanguageName = "Italian"
        Case "JapaneseLanguage"
            GetLanguageName = "Japanese"
        Case "KoreanLanguage"
            GetLanguageName = "Korean"
        Case "LatvianLanguage"
            GetLanguageName = "Latvian"
        Case "MalayLanguage"
            GetLanguageName = "Malay"
        Case "MongolianLanguage"
            GetLanguageName = "Mongolian"
        Case "NepaliLanguage"
            GetLanguageName = "Nepali"
        Case "NorwegianLanguage"
            GetLanguageName = "Norwegian"
        Case "PersianLanguage"
            GetLanguageName = "Persian"
        Case "PolishLanguage"
            GetLanguageName = "Polish"
        Case "PortugueseLanguage"
            GetLanguageName = "Portuguese"
        Case "RussianLanguage"
            GetLanguageName = "Russian"
        Case "SamoanLanguage"
            GetLanguageName = "Samoan"
        Case "SpanishLanguage"
            GetLanguageName = "Spanish"
        Case "SwedishLanguage"
            GetLanguageName = "Swedish"
        Case "TamilLanguage"
            GetLanguageName = "Tamil"
        Case "ThaiLanguage"
            GetLanguageName = "Thai"
        Case "TurkishLanguage"
            GetLanguageName = "Turkish"
        Case "UkrainianLanguage"
            GetLanguageName = "Ukrainian"
        Case "UrduLanguage"
            GetLanguageName = "Urdu"
        Case "VietnameseLanguage"
            GetLanguageName = "Vietnamese"
        Case Else
            GetLanguageName = "Unknown"
    End Select
End Function




    




