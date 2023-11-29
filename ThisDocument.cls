VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Document_Open()
    ' Call SetApiKey only if APIKey property is not defined
    Set doc = ActiveDocument
    
    If CustomDocumentPropertyExists(doc, "APIKey") Then
        api_key = doc.CustomDocumentProperties("APIKey").Value
    End If
    
    If CustomDocumentPropertyExists(doc, "fromLanguage") Then
        fromLanguage = doc.CustomDocumentProperties("fromLanguage").Value
        selectedFromLanguageIndex = GetLanguageIndex(fromLanguage)
        ribbonUI.InvalidateControl "FromLanguageDropdown"
    End If
    
    
    If CustomDocumentPropertyExists(doc, "toLanguage") Then
        toLanguage = doc.CustomDocumentProperties("toLanguage").Value
        selectedToLanguageIndex = GetLanguageIndex(toLanguage)
        ribbonUI.InvalidateControl "ToLanguageDropdown"
    End If
    
    
    If CustomDocumentPropertyExists(doc, "RephraseStyle") Then
        RephraseStyle = doc.CustomDocumentProperties("RephraseStyle").Value
        SelectedRephraseStyleIndex = GetStyleIndex(RephraseStyle)
        ribbonUI.InvalidateControl "RephraseStyle"
    End If
       
End Sub


Function GetStyleIndex(languageStyle As String) As Integer
    Select Case languageStyle
        Case "Simplify"
            GetStyleIndex = 0
        Case "Generalize"
            GetStyleIndex = 1
        Case "Informal"
            GetStyleIndex = 2
        Case "Formal"
            GetStyleIndex = 3
        Case "Professional"
            GetStyleIndex = 4
        
            
    End Select
End Function


Function GetLanguageIndex(language As String) As Integer
 Select Case language
        Case "Afrikaans"
            GetLanguageIndex = 0
        Case "Albanian"
            GetLanguageIndex = 1
        Case "Arabic"
            GetLanguageIndex = 2
        Case "Armenian"
            GetLanguageIndex = 3
        Case "Azerbaijani"
            GetLanguageIndex = 4
        Case "Bengali"
            GetLanguageIndex = 5
        Case "Bulgarian"
            GetLanguageIndex = 6
        Case "Chinese"
            GetLanguageIndex = 7
        Case "Croatian"
            GetLanguageIndex = 8
        Case "Czech"
            GetLanguageIndex = 9
        Case "Danish"
            GetLanguageIndex = 10
        Case "Dutch"
            GetLanguageIndex = 11
        Case "English"
            GetLanguageIndex = 12
        Case "Estonian"
            GetLanguageIndex = 131
        Case "Fijian"
            GetLanguageIndex = 14
        Case "Finnish"
            GetLanguageIndex = 15
        Case "French"
            GetLanguageIndex = 16
        Case "Georgian"
            GetLanguageIndex = 17
        Case "German"
            GetLanguageIndex = 18
        Case "Greek"
            GetLanguageIndex = 19
        Case "Hebrew"
            GetLanguageIndex = 20
        Case "Hindi"
            GetLanguageIndex = 21
        Case "Hungarian"
            GetLanguageIndex = 22
        Case "Icelandic"
            GetLanguageIndex = 23
        Case "Indonesian"
            GetLanguageIndex = 24
        Case "Italian"
            GetLanguageIndex = 25
        Case "Japanese"
            GetLanguageIndex = 26
        Case "Korean"
            GetLanguageIndex = 27
        Case "Latvian"
            GetLanguageIndex = 28
        Case "Malay"
            GetLanguageIndex = 29
        Case "Mongolian"
            GetLanguageIndex = 30
        Case "Nepali"
            GetLanguageIndex = 31
        Case "Norwegian"
            GetLanguageIndex = 32
        Case "Persian"
            GetLanguageIndex = 33
        Case "Polish"
            GetLanguageIndex = 34
        Case "Portuguese"
            GetLanguageIndex = 35
        Case "Russian"
            GetLanguageIndex = 36
        Case "Samoan"
            GetLanguageIndex = 37
        Case "Spanish"
            GetLanguageIndex = 38
        Case "Swedish"
            GetLanguageIndex = 39
        Case "Tamil"
            GetLanguageIndex = 40
        Case "Thai"
            GetLanguageIndex = 41
        Case "Turkish"
            GetLanguageIndex = 42
        Case "Ukrainian"
            GetLanguageIndex = 43
        Case "UrduLanguage"
            GetLanguageIndex = 44
        Case "Vietnamese"
            GetLanguageIndex = 45
    End Select
End Function
