Attribute VB_Name = "mdlShowMassMailerForm"
Public m_app As Application
Public menu As Object
Private button As Office.CommandBarButton

Const TOOLBAR_NAME As String = "XMLmassMailer"
Const APP_NAME As String = "XMLmassMailer"
Const SETTINGS = "Settings"

Sub showMassMailer()
    frmMassMailer.Show
End Sub

Public Function getToolbar() As CommandBar
    Set getToolbar = m_app.ActiveExplorer.CommandBars(TOOLBAR_NAME)
End Function

Sub makeToolbar()

    removeToolbar
    
    Set m_app = Application

    Set menu = m_app.ActiveExplorer.CommandBars.Add(TOOLBAR_NAME, msoBarTop, , True)
    menu.Protection = msoBarNoCustomize
    menu.Visible = True
    
    restoreToolbarPosition
    'restoreToolbarPosition
    Set button = menu.Controls.Add(temporary:=True)
    With button
        .Style = msoButtonIconAndCaption
        .FaceId = 29   '322 '459    '1310  '643   '1107
        .Caption = "Mass mailer"
        .OnAction = "showMassMailer"
    End With
End Sub

Sub removeToolbar()
    
    On Error Resume Next
    With getToolbar
        .Controls(1).Delete
        .Delete
    End With
    Set button = Nothing
    Set menu = Nothing

End Sub

Public Sub uninstall()
    saveToolbarPosition
    removeToolbar
    Set m_app = Nothing
End Sub

Public Sub saveToolbarPosition()
    SaveAppSetting "toolbarRowIndex", getToolbar.RowIndex
    SaveAppSetting "toolbarTop", getToolbar.Top
    SaveAppSetting "toolbarLeft", getToolbar.Left
    SaveAppSetting "toolbarPosition", getToolbar.Position
    SaveAppSetting "toolbarVisible", getToolbar.Visible
End Sub
Public Sub restoreToolbarPosition()
    getToolbar.Position = getAppSetting("toolbarPosition")
    getToolbar.RowIndex = CInt(getAppSetting("toolbarRowIndex"))
    getToolbar.Left = getAppSetting("toolbarLeft")
    getToolbar.Top = getAppSetting("toolbarTop")
    getToolbar.Visible = getAppSetting("toolbarVisible")
End Sub

Public Function getAppSetting(sKey As String, Optional valueIfNoDefaultSet As Variant = False) As Variant
    getAppSetting = GetSetting(APP_NAME, SETTINGS, sKey, getDefaultSetting(sKey, valueIfNoDefaultSet))
End Function

Public Sub SaveAppSetting(sKey As String, value As Variant)
    SaveSetting APP_NAME, SETTINGS, sKey, value
End Sub
Public Function getDefaultSetting(settingName As String, Optional notListedValue As Variant = False) As Variant

'this sets default for settings
Select Case LCase(settingName)

    Case LCase("toolbarRowIndex")
        getDefaultSetting = 5
    Case LCase("toolbarPosition")
        getDefaultSetting = 1
    Case LCase("toolbarVisible")
        getDefaultSetting = True
    Case LCase("toolbarTop")
        getDefaultSetting = 50
    Case LCase("toolbarLeft")
         getDefaultSetting = 0
    Case Else
        getDefaultSetting = notListedValue
End Select

End Function
