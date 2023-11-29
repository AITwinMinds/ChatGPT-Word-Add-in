# ChatGPT-Word-Add-in
Welcome to the ChatGPT Word Add-in—an exciting enhancement for Microsoft Office Word on your Windows desktop! 🎉🎉

![image](https://github.com/AITwinMinds/ChatGPT-Word-Add-in/assets/100919352/9e982fd1-9787-45ab-aa53-e805ad79a0c8)

Our add-in leverages the powerful ChatGPT (GPT-3.5 Turbo model) to elevate your Word experience with six key functionalities:

1- **Rephrase selection**: Give your text a fresh spin with just a click. You can select among five rephrase options:
* **Simplify**: Streamline your text for clarity.
* **Generalize**: Broaden the scope of your language.
* **Informal**: Infuse a casual and friendly tone.
* **Formal**: Elevate your writing with a polished touch.
* **Professional**: Craft your content for a business-ready presentation.


2- **Custom prompts**: Craft your own queries and get insightful responses.

3- **Email replies**: Streamline your email communication by generating quick and effective responses.

4- **Summarize text**: Condense lengthy content without losing the essence.

5- **Explain text**: Demystify complex passages with clear and concise explanations.

6- **Translate text**: Break language barriers by translating selected text seamlessly.

Experience the next level of productivity and creativity with the ChatGPT Word Add-in—your go-to tool for effortless and enhanced document handling!


# Installation
**Step 1**: Locate the Word Startup Folder
* Open File Explorer and navigate to Drive C.
* In the address bar, type %APPDATA%\Microsoft\Word and press Enter.
* Check if the "Startup" folder exists. If not, create a new folder named "Startup."

**Step 2**: Copy the Add-in File
* Copy the ChatGPT.dotm file to the "Startup" folder you just located.

**Step 3**: Open the Normal.dotm Template
* In Drive C, navigate to %AppData%\Microsoft\Templates\.
* Locate and open the file named Normal.dotm.

**Step 4**: Access the VBA Editor
* In Word, press Alt + F11 to open the Visual Basic for Applications (VBA) editor.
* In the left pane, find and select "Normal" under "Microsoft Word Objects."
* If a module already exists, double-click on it to open; otherwise, right-click on "Normal" and choose "Insert" > "Module."

**Step 5**: Write and Save the Macro
* In the code window, paste the following macro:

```{r test-python, engine='python'}
Sub AutoExec()
    AddIns.Add FileName:=Environ("AppData") & "\Microsoft\Word\Startup\ChatGPT.dotm", Install:=True
End Sub
```
Save the changes to the Normal.dotm file.

**Installation Complete**

Now, when users run Word, the ChatGPT add-in will automatically load from the specified location, enhancing their Word experience.
