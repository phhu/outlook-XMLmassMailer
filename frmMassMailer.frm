VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMassMailer 
   Caption         =   "XML mass mailer"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   OleObjectBlob   =   "frmMassMailer.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMassMailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'used to sort relative paths for attachments. See http://stackoverflow.com/questions/4613657/how-do-i-check-if-a-given-path-is-relative-or-absolute-in-vba
Private Declare Function PathCombine Lib "shlwapi.dll" Alias "PathCombineA" (ByVal szDest As String, ByVal lpszDir As String, ByVal lpszFile As String) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As _
    String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long

Private m_app As Outlook.Application

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const MAX_PATH As Long = 260

Private Declare Function GetFileAttributesA Lib "kernel32" (ByVal lpFileName As String) As Long
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10&
Private Const FILE_ATTRIBUTE_INVALID   As Long = -1&  ' = &HFFFFFFFF&

Private Const APP_NAME = "MassMailer"
Private Const SECTION = "Files"

Private sig As Word.Document
Private wordApp As Word.Application
Private sigName As String

Public sourceFile As String


Private Sub getApp()
    Set m_app = Outlook.Application

End Sub

Private Function getWordApp()
    If wordApp Is Nothing Then Set wordApp = New Word.Application
End Function


Private Sub cmdOpenDir_Click()
        ShellExecute 0, "open", getpath(Me.cmbSourceFile.value), "", "", vbNormalFocus
End Sub

Private Function getpath(ByVal inputPath As String) As String
    Dim pos As Long
    pos = InStrRev(inputPath, "\")
    If pos > 0 Then
        getpath = Left(inputPath, pos)
    Else
        getpath = getpath
    End If
End Function

Private Sub cmdOutputPath_Click()
    Me.cmbOutputPath = "c:\temp"
End Sub

Private Sub cmdPasteSource_Click()
    
    Dim v As String
    With Me.cmbSourceFile
        .value = ""
        .Paste
        v = Trim(.value)
        If Left(v, 1) = """" Then
            v = Mid(v, 2)
            If Right(v, 3) = """" & vbCrLf Then v = Mid(v, 1, Len(v) - 3)
            v = Replace(v, """""", """")
            
        End If
        .value = v
    End With
    

End Sub

Private Sub cmdPickOutlookFolder_Click()

    'Me.cmbOutlookFolder = m_app.GetNamespace("MAPI").PickFolder.FolderPath

End Sub

Private Sub cmdSourceFile_Click()
    Dim s As String, p As String
    If Me.chkUsePreviousPath.value = True Then
        p = getpath(Me.cmbSourceFile.value)
    Else
        p = CurDir
    End If
    s = getOpenFileNameFromDialog(p, _
        "XML files (*.xml)" & Chr$(0) & "*.xml" & Chr$(0) & "All files(*.*)" & Chr$(0) & "*.*", _
        "xml")
    If s <> "" Then
        Me.cmbSourceFile.value = s
    End If
End Sub

Private Sub cmdTemplate_Click()
   Dim s As String, p As String
    If Me.chkUsePreviousPath.value = True Then
        p = getpath(Me.cmbTemplate.value)
    Else
        p = CurDir
    End If
    s = getOpenFileNameFromDialog(p, _
        "OFT files (*.oft)" & Chr$(0) & "*.oft" & Chr$(0) & "All files(*.*)" & Chr$(0) & "*.*", _
        "oft")
    If s <> "" Then
        Me.cmbTemplate.value = s
    End If
End Sub

Private Sub cmdViewSourceFile_Click()
    If FileExists(Me.cmbSourceFile.value) Then
        ShellExecute 0, "open", Me.cmbSourceFile.value, "", "", vbNormalFocus
    ElseIf InStr(1, Me.cmbSourceFile.value, "<", vbTextCompare) Then
        makeFile Me.cmbSourceFile.value, Environ("temp") & "\massmailerTemp.xml"
        ShellExecute 0, "open", Environ("temp") & "\massmailerTemp.xml", "", "", vbNormalFocus
    
    End If
    
End Sub

Private Sub lblHelp_Click()
    ShellExecute 0, "open", "https://github.com/horsepress/outlook-XMLmassMailer", "", "", vbNormalFocus
End Sub

Public Function getAllSignatures() As Variant

    Dim sigs As Variant, path As String, sig As String, i As Long
    path = Environ("appdata") & "\microsoft\signatures\*.htm"
    sigs = Array()
    
    sig = Dir(path)
    push sigs, ""
    If Len(sig) > 0 Then
        Do
            push sigs, sig
            sig = Dir
        Loop Until Len(sig) < 1
    End If
    
    For i = LBound(sigs) To UBound(sigs)
        sigs(i) = Replace(sigs(i), ".htm", "")
    Next
    
    getAllSignatures = sigs
        
End Function

Private Sub UserForm_Activate()

    getApp
    populateFromSavedSettings Me.cmbSourceFile, "file", "lastFile"
    populateFromSavedSettings Me.cmbTemplate, "Template", "Template"
    Dim sigList As Variant
    sigList = getAllSignatures
    populateLst Me.cmbSignature, sigList


    Me.cmbMailitemXpath.value = GetSetting(APP_NAME, SECTION, "MailitemXpath", "/*/*")
    Me.cmbSignature.value = GetSetting(APP_NAME, SECTION, "Signature", sigList(0))
    Me.optDisplayFirst.value = GetSetting(APP_NAME, SECTION, "DisplayFirst", Me.optDisplayFirst.value)
    Me.optDisplayAll.value = GetSetting(APP_NAME, SECTION, "DisplayAll", Me.optDisplayAll.value)
    Me.optSendAll.value = GetSetting(APP_NAME, SECTION, "SendAll", Me.optSendAll.value)
    Me.chkUsePreviousPath.value = GetSetting(APP_NAME, SECTION, "PreviousPath", False)

End Sub

Sub populateFromSavedSettings(control As Object, name As String, defaultName As String)

    With control
        Dim i As Long, s As String
        For i = 1 To 10
            s = GetSetting(APP_NAME, SECTION, name & i, "")
            'If Len(s) > 0 Then .AddItem s
        Next
        .value = GetSetting(APP_NAME, SECTION, defaultName, "")
    End With

End Sub

Public Sub populateLst(lst As Object, values As Variant, Optional displayValues As Variant, Optional TwoColumns As Boolean = False)
'this populates the input combobox with the lists defined above. Default value is selected.

Dim i As Long
With lst
    .Clear
    '.MatchRequired = True
    If TwoColumns Then
        .ColumnCount = 2
        .BoundColumn = 1
        .TextColumn = 2
        .ColumnWidths = "0;"
    End If
    
    For i = LBound(values) To UBound(values)
        If TwoColumns Then
            .AddItem values(i)
            .List(i, 1) = displayValues(i)
        Else
            .AddItem values(i)
        End If
    Next
End With
End Sub

Private Sub cmdClose_Click()
    Me.Hide
    tidyUp
End Sub

Private Sub cmdGo_Click()

    If FileExists(Me.cmbSourceFile.value) Or InStr(1, Me.cmbSourceFile.value, "<", vbTextCompare) Then
        
        If Me.cmbSourceFile.value <> GetSetting(APP_NAME, SECTION, "file1", "") Then
            Dim i As Long
            For i = 10 To 1 Step -1
                SaveSetting APP_NAME, SECTION, "file" & (i + 1), GetSetting(APP_NAME, SECTION, "file" & i, "")
            Next
            SaveSetting APP_NAME, SECTION, "file1", Me.cmbSourceFile.value
        End If
        
        saveSettings
        

        Call execute
    Else
        MsgBox "File """ & Me.cmbSourceFile.value & """ does not exist. Aborting."
    End If
End Sub

Private Sub saveSettings()
        SaveSetting APP_NAME, SECTION, "Template", Me.cmbTemplate.value
        SaveSetting APP_NAME, SECTION, "Signature", Me.cmbSignature.value
        SaveSetting APP_NAME, SECTION, "LastFile", Me.cmbSourceFile.value
        SaveSetting APP_NAME, SECTION, "DisplayFirst", Me.optDisplayFirst.value
        SaveSetting APP_NAME, SECTION, "DisplayAll", Me.optDisplayAll.value
        SaveSetting APP_NAME, SECTION, "SendAll", Me.optSendAll.value
        SaveSetting APP_NAME, SECTION, "PreviousPath", Me.chkUsePreviousPath.value
        SaveSetting APP_NAME, SECTION, "MailitemXpath", Me.cmbMailitemXpath
End Sub

Private Sub execute()

Dim xmlList As New MSXML2.DOMDocument60
Dim xmlMailItems As MSXML2.IXMLDOMNodeList, xmlMailAttachments As MSXML2.IXMLDOMNodeList, mailChildNodes As MSXML2.IXMLDOMNodeList
Dim xmlMailItem As MSXML2.IXMLDOMNode, xmlMailAttachment As MSXML2.IXMLDOMNode, mailChildNode As MSXML2.IXMLDOMNode
Dim i As Outlook.MailItem
Dim attachmentFilename As String
Dim nonExistentAttachments As Variant
Dim mailItems As New Collection
Dim tmpBody As String
Dim xmlLoaded As Boolean, sourceValue As String
Dim initialHTMLbody As String, initialBody As String
xmlLoaded = False
Me.Hide
sigName = ""
'On Error GoTo doError

sourceValue = Me.cmbSourceFile.value
With xmlList
    If .Load(sourceValue) Then
        xmlLoaded = True
        sourceFile = sourceValue
    Else
        'get rid of leading ". For some reason when copying from Excel this is required.
        If Left(Trim(sourceValue), 1) = """" Then sourceValue = Mid(Trim(sourceValue), 2)
        If .loadXML(sourceValue) Then
            xmlLoaded = True
        End If
        sourceFile = ""             'we don't have a source file
    End If
    If xmlLoaded Then
    
        'check attachments exist:
        Set xmlMailAttachments = xmlList.selectNodes("//attachment")
        nonExistentAttachments = Array()
        For Each xmlMailAttachment In xmlMailAttachments
            xmlMailAttachment.Text = getAbsoluteFilePath(xmlMailAttachment.Text)
            If Not FileExists(xmlMailAttachment.Text) Then push nonExistentAttachments, xmlMailAttachment.Text
            Debug.Print xmlMailAttachment.Text
        Next
        If UBound(nonExistentAttachments) > -1 Then
            If MsgBox("Unable to find attachments: " & vbCrLf & vbCrLf & _
                dumpArray(nonExistentAttachments) & vbCrLf & vbCrLf & _
                "Continue anyway, without attaching these files?", vbYesNo) = vbNo Then
                    err.Raise -1, , "Cannot find attachments """ & dumpArray(nonExistentAttachments) & """: aborting"
            End If
        End If
        If Len(Me.cmbTemplate.value) > 0 And Not FileExists(Me.cmbTemplate.value) Then err.Raise -2, "", "Template """ & Me.cmbTemplate.value & """ does not exist"
        
        'cycle mail items
        Set xmlMailItems = xmlList.selectNodes(Me.cmbMailitemXpath.value)
        
        For Each xmlMailItem In xmlMailItems
        
            'need to be able to create item from template here
            'also need to run string replacer here
            'Outlook.CreateItemFromTemplate
            If Len(Me.cmbTemplate.value) > 0 And FileExists(Me.cmbTemplate.value) Then
                Set i = Outlook.CreateItemFromTemplate(Me.cmbTemplate.value)
            Else
                Set i = Outlook.CreateItem(olMailItem)
            End If

            initialBody = i.Body
            initialHTMLbody = i.HTMLBody

            i.To = i.To & IIf(Len(i.To) > 0, ";", "") & getMailField(xmlMailItem, "to")
            i.cc = i.cc & IIf(Len(i.cc) > 0, ";", "") & getMailField(xmlMailItem, "cc")
            i.BCC = i.BCC & IIf(Len(i.BCC) > 0, ";", "") & getMailField(xmlMailItem, "bcc")
            i.Subject = i.Subject & getMailField(xmlMailItem, "subject")
            tmpBody = getMailField(xmlMailItem, "body")
            If Len(tmpBody) > 0 Then i.Body = getMailField(xmlMailItem, "body")
            tmpBody = getMailField(xmlMailItem, "htmlbody")
            If Len(tmpBody) > 0 Then i.HTMLBody = getMailField(xmlMailItem, "htmlbody", , , True)
            i.Importance = getMailField(xmlMailItem, "@importance", olImportanceNormal)
            
            Set mailChildNodes = xmlMailItem.selectNodes("*")
            For Each mailChildNode In mailChildNodes
                On Error Resume Next

                    replaceFields i, mailChildNode.nodeName, mailChildNode.ChildNodes(0).NodeValue

                On Error GoTo 0
            Next
            
            Set xmlMailAttachments = xmlMailItem.selectNodes("attachment")
            
            For Each xmlMailAttachment In xmlMailAttachments
                attachmentFilename = xmlMailAttachment.Text
                If FileExists(attachmentFilename) Then
                    i.attachments.Add attachmentFilename
                Else
'                    If MsgBox("Unable to find attachment """ & attachmentFilename & """. Continue without attaching it?", vbYesNo) = vbNo Then
'                        err.Raise -1, , "Cannot find attachment """ & attachmentFilename & """: aborting"
'                    End If
                End If
            Next

            mailItems.Add i
                       
            If Me.optDisplayFirst Then Exit For
                
        Next
        
        Dim sendItemCount, messagesNotSent As Long
        messagesNotSent = 0
        'send or display the items
        
        'confirm send all
        If Me.optSendAll Then
            If MsgBox("Send " & mailItems.Count & " messages now?", vbYesNo, "Confirm send all") = vbNo Then
                err.Raise 5, , "Send all cancelled"
            End If
        End If
        
        Dim savePath As String, cnt As Long
        savePath = Me.cmbOutputPath
        cnt = 0
        For Each i In mailItems
            cnt = cnt + 1
            If Len(Me.cmbSignature.value) > 0 Then addSig i
            If Me.optDisplayAll Or Me.optDisplayFirst Then
                i.Display
            ElseIf Me.optSaveToFile Then
                If Not DirExists(savePath) Then err.Raise 1, , "Save path " & savePath & " does not exist."
                
                i.SaveAs savePath & "\" & CStr(cnt) & "_" & i.Subject & ".msg"
            ElseIf Me.optSaveToOutlookFolder Then

            ElseIf Me.optSendAll Then
                On Error GoTo displayMsg
                i.Send
                GoTo continueLoop
displayMsg:
                On Error GoTo 0
                messagesNotSent = messagesNotSent + 1
                i.Display
continueLoop:
                On Error GoTo 0
            End If
        Next
    Else
        MsgBox "Error in XML mailing list source file" & vbNewLine & vbNewLine & _
                    "Line: " & .parseError.Line & vbNewLine & _
                    "Pos: " & .parseError.linepos & vbNewLine & _
                    "Reason: " & .parseError.reason & vbNewLine & _
                    "Source: " & .parseError.srcText & vbNewLine & _
                    "URL: " & .parseError.URL
    End If
End With

If messagesNotSent = 1 Then MsgBox "One message could not be sent, probably because of an unresolved name." & _
"It has been displayed instead.", , "Mass mailer send notification"
If messagesNotSent > 1 Then MsgBox messagesNotSent & " messages could not be sent, probably because of unresolved names." & _
"These have been displayed instead.", , "Mass mailer send notification"

GoTo tidyUp
doError:
    MsgBox err.description, , "Error"

tidyUp:

    Set xmlList = Nothing
    Set xmlMailItems = Nothing
    Set xmlMailAttachments = Nothing
    Set xmlMailItem = Nothing
    Set xmlMailAttachment = Nothing
    Set i = Nothing
    Set mailItems = Nothing

    tidyUp
End Sub

Sub replaceFields(ByRef i As MailItem, findString As String, replacement As String)

    Const opentag As String = "###"
    Const closetag As String = "###"

    On Error Resume Next
    i.To = Replace(i.To, opentag & findString & closetag, replacement, , , vbTextCompare)
    i.cc = Replace(i.cc, opentag & findString & closetag, replacement, , , vbTextCompare)
    i.BCC = Replace(i.BCC, opentag & findString & closetag, replacement, , , vbTextCompare)
    i.Subject = Replace(i.Subject, opentag & findString & closetag, replacement, , , vbTextCompare)
    i.Body = Replace(i.Body, opentag & findString & closetag, replacement, , , vbTextCompare)
    i.HTMLBody = Replace(i.HTMLBody, opentag & findString & closetag, replacement, , , vbTextCompare)
    On Error GoTo 0
End Sub

Function getAbsoluteFilePath(inputPath As String) As String

    Dim sBuff As String * 255, basePath As String
    
    If Len(sourceFile) > 0 Then
        basePath = getpath(sourceFile)
    Else
        basePath = CurDir
    End If
    
    PathCombine sBuff, basePath, inputPath
    getAbsoluteFilePath = Left$(sBuff, InStr(1, sBuff, vbNullChar) - 1)

End Function

Sub doReplace(ByRef replaceField As String, findString As String, replacement As String)

   
    replaceField = Replace(replaceField, tag & findString & tag, replacement, , , vbTextCompare)

End Sub

Sub addSig(ByRef i As Outlook.MailItem)

    Dim insp As Outlook.Inspector
    Dim w As Word.Document
    Dim r As Word.Range
     
    Set insp = i.GetInspector
    Set w = insp.WordEditor

    'get rid of any existing signature
    On Error Resume Next
        w.Bookmarks("_MailAutoSig").Range.Delete
    On Error GoTo 0
    'this adds section to the end of the email
    Set r = w.Sections.Add.Range

    If outlookMajorVersion() >= 15 Then
        getSig (Me.cmbSignature.value)
        sig.Sections(1).Range.Copy
        Sleep 500           'sometimes it needs a pause to address clipboard...
        r.Paste
        r.NoProofing = True
        w.Bookmarks.Add "_MailAutoSig", r
    Else
        insp.Activate
        r.Select
        InsertSig2007 Me.cmbSignature.value, i
    End If
    
    If Not Me.optDisplayAll.value And Not Me.optDisplayFirst.value Then insp.Close olSave
    
    Set insp = Nothing
    Set w = Nothing
    Set r = Nothing
    
End Sub


Sub getSig(sigFilename As String)

    getWordApp
    If sig Is Nothing Or sigFilename <> sigName Then
        sigName = sigFilename
        Set sig = wordApp.Documents.Open(filename:=Environ("appdata") & "\Microsoft\Signatures\" & sigFilename & ".htm", ReadOnly:=True, Visible:=False)
    End If
    
End Sub

Sub InsertSig2007(strSigName As String, objMailItem As MailItem)
    Dim objItem As Object
    Dim objInsp As Outlook.Inspector
    ' requires a project reference to the
    ' Microsoft Office library
    Dim objCBP As Office.CommandBarPopup
    Dim objCBP2 As Office.CommandBarPopup
    Dim objCBB As Office.CommandBarButton
    Dim colCBControls As Office.CommandBarControls
    On Error Resume Next
     
    Set objInsp = objMailItem.GetInspector
    If Not objInsp Is Nothing Then
        Set objItem = objInsp.CurrentItem
        If objItem.Class = olMail Then
            ' get Insert menu
            Set objCBP = objInsp.CommandBars.ActiveMenuBar.FindControl(, 30005)
            ' get Signature submenu
            Set objCBP2 = objCBP.CommandBar.FindControl(, 5608)
            If Not objCBP2 Is Nothing Then
                Set colCBControls = objCBP2.Controls
                For Each objCBB In colCBControls
                    If objCBB.Caption = strSigName Then
                        objCBB.execute ' **** see remarks
                        Exit For
                    End If
                Next
            End If
        End If
    End If
     
    Set objInsp = Nothing
    Set objItem = Nothing
    Set colCBControls = Nothing
    Set objCBB = Nothing
    Set objCBP = Nothing
    Set objCBP2 = Nothing
End Sub



Sub InsertSig(strSigName As String, Optional objInsp As Outlook.Inspector)
    Dim objItem As Object
    'Dim objinsp As Outlook.Inspector
    ' requires a project reference to the
    ' Microsoft Word library
    Dim objDoc As Word.Document
    Dim objSel As Word.Selection
    ' requires a project reference to the
    ' Microsoft Office library
    Dim objCB As Office.CommandBar
    Dim objCBP As Office.CommandBarPopup
    Dim objCBB As Office.CommandBarButton
    Dim colCBControls As Office.CommandBarControls
    On Error Resume Next
     
    If objInsp Is Nothing Then Set objInsp = Application.ActiveInspector
    'Set objinsp = Application.ActiveInspector
    If Not objInsp Is Nothing Then
        Set objItem = objInsp.CurrentItem
        If objItem.Class = olMail Then  ' editor is WordMail
            If objInsp.EditorType = olEditorWord Then
                ' next statement will trigger security prompt
                ' in Outlook 2002 SP3
                Set objDoc = objInsp.WordEditor
                Set objSel = objDoc.Application.Selection
                If objDoc.Bookmarks("_MailAutoSig") Is Nothing Then
                    objDoc.Bookmarks.Add Range:=objSel.Range, name:="_MailAutoSig"
                End If
                objSel.GoTo What:=wdGoToBookmark, name:="_MailAutoSig"
                
                Set objCB = objDoc.CommandBars("AutoSignature Popup")
                If Not objCB Is Nothing Then
                    Set colCBControls = objCB.Controls
                End If
            Else ' editor is not WordMail
                ' get the Insert | Signature submenu
                Set objCBP = Application.ActiveInspector.CommandBars.FindControl(, 31145)
                If Not objCBP Is Nothing Then
                    Set colCBControls = objCBP.Controls
                End If
            End If
        End If
        If Not colCBControls Is Nothing Then
            For Each objCBB In colCBControls
                If objCBB.Caption = strSigName Then
                    objCBB.execute ' **** see remarks
                    Exit For
                End If
            Next
        End If
    End If
     
    Set objInsp = Nothing
    Set objItem = Nothing
    Set objDoc = Nothing
    Set objSel = Nothing
    Set objCB = Nothing
    Set objCBB = Nothing
End Sub

Private Function getMailField(node As MSXML2.IXMLDOMNode, xpath As String, _
    Optional default As String = "", Optional asXML As Boolean = False, _
    Optional asXMLifHasChildNodes As Boolean = False) As String
    
    Dim n As MSXML2.IXMLDOMNode
    On Error Resume Next
    getMailField = default
    '//mailitem/*[self::to or self::subject]
    If asXMLifHasChildNodes Then
        asXML = node.selectSingleNode(xpath).FirstChild.HasChildNodes
    End If
    
    
    If asXML Then
        getMailField = ""
        For Each n In node.selectSingleNode(xpath).ChildNodes
            getMailField = getMailField & n.XML
        Next
    Else
        getMailField = node.selectSingleNode(xpath).Text
    End If
    
    Set n = Nothing
    
End Function


Private Function FileExists(ByVal sPathName As String) As Boolean
' -------------------------------------------------------------------
' Funktion: Prüft Existenz von Datei, schneller als Dir
'   Since we only want to return TRUE for a file in this case, we only need to
'   check for a set '1' in the directory flag position.
'
' Parameter: keine
' Rückgabewerte: wahr, wenn vorhanden
' Aufgerufene Prozeduren: GetFileAttributesA
' letzte Änderung: 26.05.2002
' -------------------------------------------------------------------
    Dim attr As Long

    attr = GetFileAttributesA(sPathName)

    ' The directory bit is set if the path is a directory
    ' or if it does not exist (in which case attr will be -1,
    ' which includes a set directory bit).
    FileExists = ((attr And FILE_ATTRIBUTE_DIRECTORY) = 0)
End Function


Public Function getOpenFileNameFromDialog( _
    Optional initialDir As String = "C:\", _
    Optional filterUseChrDollarZeroAsSep As String = "", _
    Optional defaultExtentionNoDot As String _
    ) As String

Dim filter As String
filter = filterUseChrDollarZeroAsSep

If filter = "" Then filter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)

Dim OFName As OPENFILENAME
OFName.lStructSize = Len(OFName)
'Set the parent window
'OFName.hwndOwner = Application.hwnd
'Set the application's instance
'OFName.hInstance = Application.hInstance
'Select a filter
OFName.lpstrFilter = filter     '"Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
'create a buffer for the file
OFName.lpstrFile = Space$(254)
'set the maximum length of a returned file
OFName.nMaxFile = 255
'Create a buffer for the file title
OFName.lpstrFileTitle = Space$(254)
'Set the maximum length of a returned file title
OFName.nMaxFileTitle = 255
'Set the initial directory
OFName.lpstrInitialDir = initialDir
'Set the title
OFName.lpstrTitle = "Open file"
'No flags
OFName.flags = 0
OFName.lpstrDefExt = defaultExtentionNoDot
'Show the 'Open File'-dialog
If GetOpenFileName(OFName) Then
    getOpenFileNameFromDialog = Trim$(OFName.lpstrFile)
Else
    getOpenFileNameFromDialog = vbNullString
End If

End Function

'********************** FILE FUNCTIONS *****************

Sub makeFile(contents As String, Optional filename As String = "C:\temp\vbdata.txt")

On Error GoTo doError

DirExistsCreate getpath(filename)

Dim FileNum As Integer
    FileNum = FreeFile ' next file number
    Open filename For Output As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, contents ' write information to the text file
    Close #FileNum ' close the file
    
Exit Sub
doError:
    'errMsgbox err, "MakeFile failed"
    
End Sub
Public Function DirExists(ByVal sPathName As String, Optional stripFileName As Boolean = False) As Boolean
' -------------------------------------------------------------------
' Funktion: Prüft, ob Verzeichnis existiert
' Parameter: Path, trailing \ does not matter
' Rückgabewerte: wahr, wenn existent
' Aufgerufene Prozeduren: GetFileAttributesA
' letzte Änderung: 26.05.2002
' -------------------------------------------------------------------
    Dim attr As Long, lastSlash As Long
    
    If stripFileName Then
        lastSlash = InStrRev(sPathName, "\")
        If lastSlash > 1 Then sPathName = Left(sPathName, lastSlash - 1)
        'sPathName = regExpReplace(sPathName, "(.*)\\.*?$", "$1")
    End If
    attr = GetFileAttributesA(sPathName)

    DirExists = Not (attr = FILE_ATTRIBUTE_INVALID)

End Function
Public Function DirExistsCreate(ByVal path$, Optional errMsg As Boolean = True) As Boolean
' -------------------------------------------------------------------
' Function: creates path of any depth
' returns: true, if exists or could be created
' Released: 06-MAY-2005
' -------------------------------------------------------------------

On Error Resume Next

    If DirExists(path) Then DirExistsCreate = True: Exit Function
    
    'Path = FormatPath(Path)
    If Not DirExists(path) Then SHCreateDirectoryEx 0, path, ByVal 0&
    If DirExists(path) Then
        DirExistsCreate = True
    ElseIf errMsg Then
        'MsgBox "Der Path: " & Path & vbCrLf & "konnte nicht angelegt werden. Ändern Sie den Path.", vbExclamation, App.Title
        MsgBox "The directory: " & path & vbCrLf & "could not be created.", vbExclamation
    End If

End Function

' ****************** ARRAY FUNCTIONS*******************

Public Sub push(ByRef inputArray As Variant, value As Variant)

Dim ub As Integer
ub = UBound(inputArray)

ReDim Preserve inputArray(ub + 1)
inputArray(ub + 1) = value

End Sub

Public Function dumpArray(sourceArray As Variant, Optional printIndex As Boolean = True, _
Optional beforeText As String, Optional betweentext As String = ": ", _
Optional aftertext As String = vbCrLf, Optional reverseOrder As Boolean = False) As String

If Not IsArray(sourceArray) Then
    dumpArray = "[not an array]"
    Exit Function
End If
Dim a As Long, lb As Long, ub As Long, stp As Long

If reverseOrder Then
    lb = UBound(sourceArray)
    ub = LBound(sourceArray)
    stp = -1
Else
    ub = UBound(sourceArray)
    lb = LBound(sourceArray)
    stp = 1
End If

If printIndex Then
    For a = lb To ub Step stp
        dumpArray = dumpArray & beforeText & a & betweentext & sourceArray(a) & aftertext
    Next
Else
    For a = lb To ub Step stp
        dumpArray = dumpArray & beforeText & sourceArray(a) & aftertext
    Next
End If

End Function


Public Function GetFolder(strFolderPath As String) As MAPIFolder
  ' strFolderPath needs to be something like
  '   "Public Folders\All Public Folders\Company\Sales" or
  '   "Personal Folders\Inbox\My Folder"

  Dim objApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Dim colFolders As Outlook.Folders
  Dim objFolder As Outlook.MAPIFolder
  Dim arrFolders() As String
  Dim i As Long
  On Error Resume Next

  strFolderPath = Replace(strFolderPath, "/", "\")
  arrFolders() = Split(strFolderPath, "\")
  Set objApp = Application
  Set objNS = objApp.GetNamespace("MAPI")
  Set objFolder = objNS.Folders.Item(arrFolders(0))
  If Not objFolder Is Nothing Then
    For i = 1 To UBound(arrFolders)
      Set colFolders = objFolder.Folders
      Set objFolder = Nothing
      Set objFolder = colFolders.Item(arrFolders(i))
      If objFolder Is Nothing Then
        Exit For
      End If
    Next
  End If

  Set GetFolder = objFolder
  Set colFolders = Nothing
  Set objNS = Nothing
  Set objApp = Nothing
End Function

Function outlookMajorVersion() As Long
    On Error Resume Next
    outlookMajorVersion = 15
    outlookMajorVersion = CLng(Left(Outlook.Version, InStr(1, Outlook.Version, ".") - 1))

End Function

Private Sub UserForm_Terminate()
    'on error resume next
    tidyUp
End Sub

Function ifNotBlank(value, checkValue) As String

    ifNotBlank = IIf(checkValue, checkValue, value)
    


End Function

Sub tidyUp()

    If Not wordApp Is Nothing Then wordApp.Quit savechanges:=False
    Set sig = Nothing
    Set wordApp = Nothing

End Sub
