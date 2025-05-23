Sub Main()
	
	Dim userPlane As String
	userPlane = Parameter("UserPlane1")
	Dim namePlane As String
	
	If userPlane Is Nothing Then
	    MsgBox("Invalid selection. Please enter a value for the Sketch")
		Exit Sub
	End If
	
	Dim POWidth1 As Double =  Parameter("UserPOWidth")
	Dim POHeight1 As Double =  Parameter("UserPOHeight")
	Dim KSWidth1  As Double =  Parameter("UserKSWidth")
	Dim KSHeight1  As Double =  Parameter("UserKSHeight")
	Dim POfromFloor1 As Double  = Parameter("UserPOFloor")
	Dim tempEdgeLength As Double = Parameter("UserPO1Edge")
	
	Dim temploop As Boolean = True
	Dim temploop1 As Boolean = True
	Dim sum As Double = 0
	
	Parameter("POWidth") = POWidth1
	Parameter("POHeight")= POHeight1
	Parameter("KSWidth") = KSWidth1 
	Parameter("KSHeight")= KSHeight1 
	Parameter("POfromFloor") = POfromFloor1
	Parameter("OldPOEdge1") = tempEdgeLength
	
	If userPlane = "Front" Then
		Parameter("POfromEdgeFront") = tempEdgeLength
		Parameter("OldSketch1") = "Sketch Front"
		Parameter("OldSketchPlane1") = "Front"
		namePlane = "Front Plane"
	Else If userPlane = "Right" Then
		Parameter("POfromEdgeRight") = tempEdgeLength
		Parameter("OldSketch1") = "Sketch Right"
		Parameter("OldSketchPlane1") = "Right"
		namePlane = "Right Plane"
	Else
		MsgBox("There was an error calculating the PO 1 Centerline from Sketch 1")
	End If
	
	TurnOffSketches()

	Dim temp As WorkPlane
	temp = findPlane(namePlane)
	ActivateSketch(temp)
	
End Sub


Function findPlane(namePlane As String) As WorkPlane

	 'Get the active drawing Document
    Dim workPlaneOjb As WorkPlane
    For Each workPlaneOjb In ThisApplication.ActiveDocument.ComponentDefinition.WorkPlanes
        If workPlaneOjb.Name = namePlane Then
			findPlane = workPlaneOjb
           Exit For
        End If
    Next
	
End Function



Function ActivateSketch(namePlane As WorkPlane) As Sketch
	
	Dim tempSketch As String
	If namePlane.Name = "Front Plane" Then
		tempSketch = "Sketch Front"
	Else If namePlane.Name = "Right Plane" Then
		tempSketch = "Sketch Right"
	Else
		MsgBox("Error was found in ActivateSketch Function")
		Exit Function
	End If
	
	Parameter("SketchName1") = tempSketch
	Dim oADoc As AssemblyDocument = ThisDoc.Document
	Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition
	Dim findSketchName As PlanarSketch = oADef.Sketches.Item(tempSketch)
	
	
	
	findSketchName.Visible = True
	findSketchName.Edit
	findSketchName.ExitEdit
	ActivateSketch = findSketchName
	Return ActivateSketch
	
End Function

Function TurnOffSketches()
	
	Dim oADoc As AssemblyDocument = ThisDoc.Document
	Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition
	Dim findSketchName As PlanarSketch = oADef.Sketches.Item("Sketch Right")
	findSketchName.Visible = False
	Dim findSketchName1 As PlanarSketch = oADef.Sketches.Item("Sketch Front")
	findSketchName1.Visible = False
	
End Function

