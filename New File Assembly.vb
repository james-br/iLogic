AddReference "Autodesk.iLogic.Core.dll"
AddReference "Autodesk.iLogic.UiBuilderCore.dll"
Imports iLogicCore = Autodesk.iLogic.Core

Sub Main()

	If ThisDoc.Document.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
		MsgBox("An Assembly Document must be active for this rule to work.", vbCritical, "Assembly Check")
		Exit Sub
	End If
	
	Dim oADoc As AssemblyDocument = ThisDoc.Document
	Dim oRefDocs As DocumentsEnumerator = oADoc.AllReferencedDocuments
	If oRefDocs.Count = 0 Then
		MsgBox("No referenced documents found.", vbExclamation, "Reference Check")
		Exit Sub
	End If

	' Prompt for SaveAs location
	Dim sNewFullFileName As String = UseSaveAsDialog()
	If sNewFullFileName = "" Then
		
		MsgBox("SaveAs cancelled or no filename provided.", vbInformation, "Cancelled")
		Exit Sub
	End If

	
	Try
		oADoc.SaveAs(sNewFullFileName, True)
		Parameter("ExitSave") = False
		MsgBox("SaveAs Success!", vbInformation, "New File")
		
	Catch ex As Exception
		MsgBox("SaveAs Failed: " & ex.Message, vbCritical, "Error")
	End Try
	
	Try
		ThisApplication.Documents.Open(sNewFullFileName, True)
		MsgBox("New file opened: " & vbCrLf & sNewFullFileName, vbInformation, "Success")
	Catch ex As Exception
		MsgBox("Failed to open new file: " & ex.Message, vbCritical, "Error")
	End Try
	

End Sub

Function UseSaveAsDialog() As String
	Dim oFileDialog As Inventor.FileDialog = Nothing
	ThisApplication.CreateFileDialog(oFileDialog)

	oFileDialog.DialogTitle = "Specify New Name & Location For Copied Assembly"
	oFileDialog.InitialDirectory = "C:\Vault Works\Designs\Units\Ovens"
	oFileDialog.Filter = "Autodesk Inventor Assemblies (*.iam)|*.iam"
	oFileDialog.FileName = iProperties.Value("Project", "Part Number")
	oFileDialog.MultiSelectEnabled = False
	oFileDialog.OptionsEnabled = False
	oFileDialog.InsertMode = False
	oFileDialog.CancelError = True

	On Error Resume Next
	oFileDialog.ShowSave()
	Parameter("FileTitle") = oFileDialog.FileName
	If Err.Number <> 0 Then
		UseSaveAsDialog = ""
		Exit Function
	End If
	On Error GoTo 0

	UseSaveAsDialog = oFileDialog.FileName
End Function


