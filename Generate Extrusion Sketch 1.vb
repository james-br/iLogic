Sub Main()
	Dim sketchName1 As String = SketchName1
	
	If POAssembly1On = True Then
		iLogicVb.RunRule("Rule 33: Suppress Flash 1 Assembly")
		ExtrusionSketch(GenerateDrawing(sketchName1))
		iLogicVb.RunRule("Rule 35: Unsuppress Flash 1 Assembly")
	Else
		ExtrusionSketch(GenerateDrawing(sketchName1))
		Parameter("OldSketch1") = sketchName1
	End If

End Sub



Function GenerateDrawing(tempSketch As String) As Sketch

	Dim oADoc As AssemblyDocument = ThisDoc.Document
	Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition
	Dim sketch4 As PlanarSketch = oADef.Sketches.Item(tempSketch)
	sketch4.Edit
	sketch4.ExitEdit
	GenerateDrawing = sketch4
	Return GenerateDrawing
	
End Function

Function ExtrusionSketch(foundSketch As PlanarSketch)
	
	If TypeName(foundSketch) = "PlanarSketch" Then
			' Create a profile from the sketch geometry
		    Dim profile As Profile
		    profile = foundSketch.Profiles.AddForSolid
		
		    ' Create the extrusion feature
		    Dim extrudeDef As ExtrudeDefinition
		    extrudeDef = ThisApplication.ActiveDocument.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(profile, PartFeatureOperationEnum.kCutOperation)
			
		    ' Set extrusion parameters			
		    extrudeDef.SetDistanceExtent (20, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
			
		    ' Create the extrusion
		    Dim extrude As ExtrudeFeature
			
		    extrude = ThisApplication.ActiveDocument.ComponentDefinition.Features.ExtrudeFeatures.Add(extrudeDef)
			extrude.Name = "Opening Extrusion 1"
			
			
	Else 
		MsgBox("Parameter passed was not a Sketch")
	End If
	   

End Function

