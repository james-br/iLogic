iLogicVb.RunRule("Rule 33: Suppress Flash 1 Assembly")
Dim tempSketch As String = OldSketch1
If (tempSketch Is Nothing) Then
    MsgBox("SketchName1 is not defined. Exiting.", vbCritical, "Error")
    Exit Sub
End If

Dim oADoc As AssemblyDocument = ThisDoc.Document
Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition
Dim oSketch2 As PlanarSketch = oADef.Sketches.Item(tempSketch)
Dim oExtFeats As ExtrudeFeatures = oADef.Features.ExtrudeFeatures
If oExtFeats.Count = 0 Then Exit Sub
Dim oExtFeat As ExtrudeFeature = Nothing
For Each oEF As ExtrudeFeature In oExtFeats
	If oEF.Definition.Profile.Parent Is oSketch2 Then
		oEF.Delete(True, True, True) 'try to preserve everything else associated with it
		
		
	End If
Next
Parameter("POAssembly1On") = False
