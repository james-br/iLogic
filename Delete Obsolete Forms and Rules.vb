AddReference "Autodesk.iLogic.Core.dll"
AddReference "Autodesk.iLogic.UiBuilderCore.dll"
Imports iLogicCore = Autodesk.iLogic.Core



iLogicVb.RunRule("Remove Inactive")
iLogicVb.RunRule("Delete Inactive")
Dim oDoc = ThisApplication.ActiveDocument 

iCount = 0
Dim oUIatts As New iLogicCore.UiBuilderStorage.UiAttributeStorage(oDoc)
For Each oName In oUIatts.FormNames
  '  Dim oFormsSpecs = oUIatts.LoadFormSpecification(oName)
    iCount = iCount +1                     
Next

If iCount = 0 Then
	MsgBox("No intneral iLogic forms found.",,"iLogic")
	Return 'exit rule
End If

RUSure = MessageBox.Show(iCount & " internal iLogic forms found." _
	& vbLf & "This will delete all of these internal forms." _
	& vbLf & "Are you sure you want to continue?", "iLogic", _
		MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)

If RUSure = vbNo Then Return 'exit rule
	
'delete all internal forms
Dim Attset As Inventor.AttributeSet 

For Each Attset In oDoc.AttributeSets
	If Attset.Name Like "iLogicInternalUi*" Then
		Attset.Delete 
	End If 
Next

'blink the browser to clear memory
For Each oDockableWindow As Inventor.DockableWindow _
		In ThisApplication.UserInterfaceManager.DockableWindows
	If oDockableWindow.InternalName = "ilogic.treeeditor" Then
		oDockableWindow.Visible = False
		oDockableWindow.Visible = True
	End If
Next



doc = ThisApplication.ActiveDocument
Logic = iLogicVb.Automation
iLogicVb.Automation.DeleteAllRules(doc)        



