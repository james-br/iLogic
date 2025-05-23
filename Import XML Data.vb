'XML IMPORT rule
'Code by @ClintBrown3d originally posted at https://clintbrown.co.uk/using-xml-to-drive-ilogic-configurations
'Special thanks to Curtis for sharing his File save code, which I have adapted for opeing XML's
'https://inventortrenches.blogspot.com/2012/10/ilogic-adding-save-as-dialog-box.html
oDoc = ThisDoc.Document
Dim oFileDlg As Inventor.FileDialog = Nothing
InventorVb.Application.CreateFileDialog(oFileDlg)

oFileDlg.Filter = "XML Files (*.xml)|*.xml"
oFileDlg.DialogTitle = "Import XML"
'set the directory to open the dialog at
oFileDlg.InitialDirectory = "C:\Vault Works\Designs\Units\Ovens\XML Configurations\"
'oFileDlg.InitialDirectory = "C:\Vault Works\Designs\Units\Environmental Rooms\XML Configurations\"
oFileDlg.ShowOpen()'Show File Open Dialogue
ClintBrown3D = oFileDlg.FileName
If ClintBrown3D = "" Then: Return: End If

'Open the selected file
On Error GoTo ClintsErrorTrapper 'to handle an exit without selecting a file
'	ThisDoc.Launch(ClintBrown3D)
	iLogicVb.Automation.ParametersXmlLoad(ThisDoc.Document, ClintBrown3D)

'Update the file
iLogicVb.UpdateWhenDone = True

iLogicVb.RunRule("Update Model")

Return

ClintsErrorTrapper:
MsgBox("error")
