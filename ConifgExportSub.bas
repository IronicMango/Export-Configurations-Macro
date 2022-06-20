Attribute VB_Name = "ConifgExportSub"
'Written by Ironic Mango Designs
'Released under MIT License

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Option Explicit

Function Export_Configurations(swModel As ModelDoc2, customSaveLocation As Boolean, configSelection As Integer) As String

'For getting configurations and their names
Dim vConfigNameArr As Variant
Dim vConfigName As Variant
Dim sConfigName As String
Dim boolstatus As Boolean
Dim parentConfig As Configuration
Dim ConfigMgr As ConfigurationManager
Dim vChildren As Variant
Dim derivedConfig As Configuration
Dim isDerived As Boolean

'For getting a save location
Dim currentFileLocation As String
Dim configFolder As String
Dim configFolderIndex As Integer
Dim folderExists As String
Dim fileOverride As String

'General app and counting variables
Dim swapp As Object
Dim i As Integer

'For saving the file
Dim swExt As ModelDocExtension
Dim isSaved As Boolean
Dim saveErrors As Long
Dim saveWarnings As Long
Dim exportData As Object
Dim setStepExport As Boolean

'Error Reporting
Dim errorMessage As String

Set swapp = Application.SldWorks

try_:
    On Error GoTo catch_
    
    vConfigNameArr = swModel.GetConfigurationNames
    
    If customSaveLocation = False Then
        currentFileLocation = swModel.GetPathName
        configFolderIndex = InStrRev(currentFileLocation, "\")
        configFolder = Left(currentFileLocation, configFolderIndex) & "configExport"
        folderExists = Dir(configFolder, vbDirectory)
        If folderExists = "" Then
            MkDir configFolder
            Debug.Print "No"
        Else
            fileOverride = MsgBox("Default export folder already exists. Do you want to overrite files?", vbYesNo + vbQuestion)
            If fileOverride = vbNo Then
                GoTo finally_
            End If
        End If
    Else
        'Determine file path for Dropbox
        currentFileLocation = swModel.GetPathName
        configFolderIndex = InStr(currentFileLocation, "Luke Lamp Co. Dropbox\") + 21
        'Get folder starting in dropbox
        configFolder = BrowseFolder(Left(currentFileLocation, configFolderIndex))
        If configFolder = "" Then
            GoTo finally_
        End If
    End If
    
    
    Set swExt = swModel.Extension
    
    setStepExport = swapp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swStepAP, 214)
    Debug.Print setStepExport
    
    'For saving Parent Configs only
    If configSelection = 0 Then
        vConfigNameArr = swModel.GetConfigurationNames
        For Each vConfigName In vConfigNameArr
            sConfigName = vConfigName
            'ensure configuration is not a derived config
            boolstatus = swModel.ShowConfiguration2(sConfigName)
            Set parentConfig = swModel.GetActiveConfiguration
            isDerived = parentConfig.isDerived
            If isDerived = False Then
                'Save As
                isSaved = swExt.SaveAs(configFolder & "\" & sConfigName & ".STEP", 0, 1, exportData, saveErrors, saveWarnings)
                If isSaved = False Then
                    Call RecordErrors(errorMessage, sConfigName, saveErrors, saveWarnings)
                End If
            End If
        Next vConfigName
    'For saving just derived configs of the selected parent
    ElseIf configSelection = 1 Then
        'Checks if parent config is selected
        Set parentConfig = swModel.GetActiveConfiguration
        isDerived = parentConfig.isDerived
        If isDerived Then
            MsgBox "Selected Configuration is a derived Configuration. Please select the parent config, and try again.", vbOKOnly + vbInformation
            GoTo finally_
        End If
        'check for child configs
        If parentConfig.GetChildrenCount() = 0 Then
            MsgBox "Selected Configuration has no derived Configurations. Please select a different parent config, or change function parameters.", vbOKOnly + vbInformation
            GoTo finally_
        End If
        'get config names
        vChildren = parentConfig.GetChildren
        ReDim vConfigNameArr(UBound(vChildren))
        For i = 0 To UBound(vChildren)
            Set derivedConfig = vChildren(i)
            vConfigNameArr(i) = derivedConfig.Name
        Next i
        
        'save as step files
        For Each vConfigName In vConfigNameArr
            sConfigName = vConfigName
            'ensure configuration is not an exploded configuration
            If InStr(sConfigName, "Exploded") = 0 Then
                boolstatus = swModel.ShowConfiguration2(sConfigName)
                'Save As
                isSaved = swExt.SaveAs(configFolder & "\" & sConfigName & ".STEP", 0, 1, exportData, saveErrors, saveWarnings)
                If isSaved = False Then
                    Call RecordErrors(errorMessage, sConfigName, saveErrors, saveWarnings)
                End If
            End If
        Next vConfigName
    'For saving all Configurations
    ElseIf configSelection = 2 Then
        vConfigNameArr = swModel.GetConfigurationNames
        For Each vConfigName In vConfigNameArr
            sConfigName = vConfigName
            boolstatus = swModel.ShowConfiguration2(sConfigName)
            'Save As
            isSaved = swExt.SaveAs(configFolder & "\" & sConfigName & ".STEP", 0, 1, exportData, saveErrors, saveWarnings)
            If isSaved = False Then
                Call RecordErrors(errorMessage, sConfigName, saveErrors, saveWarnings)
            End If
        Next vConfigName
    Else
        MsgBox "configSelection parameter is invalid." & vbCrLf & "0 = Save only Parent configs" & vbCrLf & "1 = Save derived configs of selected parent config" & vbCrLf & "2 = Save all parent and derived configurations", vbOKOnly + vbCritical
        GoTo finally_
    End If
    
    GoTo finally_
    
catch_:
    swapp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk

finally_:

    If errorMessage <> "" Then
        Export_Configurations = errorMessage
    Else
        Export_Configurations = "Completed without errors."
    End If

End Function

Function BrowseFolder(Optional openingFolder As String) As String

Dim objShell As New Shell32.Shell
Dim objFolder As Shell32.Folder

Set objFolder = objShell.BrowseForFolder(0, "Select Folder", 1, openingFolder)
If Not objFolder Is Nothing Then
    If objFolder = "Desktop" Then
        BrowseFolder = Environ("USERPROFILE") & "\Desktop"
    Else
        BrowseFolder = objFolder.Items.Item.Path
    End If
ElseIf objFolder Is Nothing Then
    MsgBox "Export Canceled", vbOKOnly + vbInformation
    Exit Function
End If

End Function

Sub RecordErrors(errorMessage As String, sConfigName As String, saveErrors As Long, saveWarnings As Long)
    
If errorMessage = "" Then
    errorMessage = "The following Configurations could not be saved:" & vbCrLf
End If

errorMessage = errorMessage & sConfigName & vbCrLf & "    Error code: " & saveErrors & vbCrLf & "    Warning Code: " & saveWarnings & vbCrLf

End Sub
