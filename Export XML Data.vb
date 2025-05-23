' Initialize variables
Dim oDirectory As String
Dim oXMLname As String
Dim oSaverXML As String
Dim i As Integer

' Define the directory path and XML file name
oDirectory = "C:\Vault Works\Designs\Units\Ovens\XML Configurations\"
oXMLname = Parameter("JobName") & ".xml"
oSaverXML = oDirectory & oXMLname

' Debugging output
MsgBox ("Saving to path: " & oSaverXML, , "Debug Information")

' Create the directory if it doesn't exist
On Error Resume Next
MkDir(oDirectory)
If Err.Number <> 0 Then
    'MsgBox ("Error creating directory: " & Err.Description, , "Error")
    Err.Clear
End If
On Error GoTo 0

' Check if the file exists and add a sequential number if it does
If System.IO.File.Exists(oSaverXML) Then
    i = 1
    Do While System.IO.File.Exists(oSaverXML)
        oXMLname = Parameter("JobName") & "-" & i & ".xml"
        oSaverXML = oDirectory & oXMLname
        i = i + 1
    Loop
End If

' Debugging output
MsgBox ("Final file path: " & oSaverXML, , "Debug Information")

' Try exporting the parameters to XML with KeysOnly set to True
On Error Resume Next
iLogicVb.Automation.ParametersXmlSave(ThisDoc.Document, oSaverXML, True)
If Err.Number <> 0 Then
    MsgBox ("Error saving XML: " & Err.Description, , "Error")
    Err.Clear
End If
On Error GoTo 0

' Confirm the operation result
If System.IO.File.Exists(oSaverXML) Then
    MsgBox ("The Oven parameters have been saved successfully at " & vbCrLf & vbCrLf & oSaverXML, , "Successfully Exported")
Else
    MsgBox ("Failed to save the XML file. Please check the folder and file path.", , "Export Failed")
End If
