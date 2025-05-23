Sub Main()
	
	If ThisDoc.FileName(False) = "Oven Configurator V4 "
		MsgBox("Please save a copy of the file before deleting inactive components.", ,"DO NOT DELETE")
		iLogicVb.RunRule("Save As")
		Exit Sub
		Else
		Dim response As MsgBoxResult = MsgBox("Deleting inactive files will prevent the configurator from running correctly. Only do this to the final design", MsgBoxStyle.OkCancel, "WARNING!")
		If Not response = MsgBoxResult.Ok Then Exit Sub
	End If
	
	RemovePatternedItems()
	Dim doc As AssemblyDocument = ThisDoc.Document
	Dim compDef As ComponentDefinition = doc.ComponentDefinition
	Dim inactiveComponents As New List(Of ComponentOccurrence)
	Dim deletionErrors As New List(Of String)
	
	' Iterate through all occurrences in the assembly
	For Each compOcc As ComponentOccurrence In compDef.Occurrences
	    ' Check if the component is suppressed (inactive)
	    If compOcc.Suppressed Then
	        inactiveComponents.Add(compOcc)
	    End If
	Next
	
	' If there are inactive components, prompt the user for deletion
	If inactiveComponents.Count > 0 Then
	    Dim message As String = "The following components are inactive:" & vbCrLf & vbCrLf
	    Dim maxComponentsToShow As Integer = 5 ' Maximum number of components to show
	    Dim componentsToShow As Integer = Math.Min(maxComponentsToShow, inactiveComponents.Count)
	    
	    ' Add the first few components to the message
	    For i As Integer = 0 To componentsToShow - 1
	        message &= inactiveComponents(i).Name & vbCrLf
	    Next
	
	    ' If there are more components than the maximum, add an ellipsis
	    If inactiveComponents.Count > maxComponentsToShow Then
	        message &= "..." & vbCrLf
	    End If
	
	    Dim response As MsgBoxResult = MsgBox(message & vbCrLf & "Do you want to delete these components?", MsgBoxStyle.YesNo, "Delete Inactive Components")
	 
	    ' If the user confirms deletion, delete the inactive components
	    If response = MsgBoxResult.Yes Then
	        For Each inactiveComp As ComponentOccurrence In inactiveComponents
	            Try
	                inactiveComp.Delete()
	            Catch ex As Exception
	                deletionErrors.Add(inactiveComp.Name)
	            End Try
	        Next
	        If deletionErrors.Count > 0 Then
	            Dim errorMessages As String = "The following components could not be deleted. If they are part of a pattern, please make them independent and delete them individually:" & vbCrLf & vbCrLf
	            For Each errorMsg As String In deletionErrors
	                errorMessages &= errorMsg & vbCrLf
	            Next
	            MsgBox(errorMessages, MsgBoxStyle.Critical, "Deletion Errors")
	        Else
	            MsgBox("Inactive components deleted successfully.", MsgBoxStyle.Information, "Success")
	        End If
	    Else
	        MsgBox("Operation cancelled by user.", MsgBoxStyle.Information, "Cancelled")
	    End If
	Else
	    MsgBox("No inactive components found.", MsgBoxStyle.Information, "Info")
	End If
	
End Sub


Sub RemovePatternedItems()
	
    Dim oDoc As Document
    oDoc = ThisDoc.Document
    
    Dim oAsmCD As AssemblyComponentDefinition
    oAsmCD = oDoc.ComponentDefinition
    
    Dim oObjToDelete As Object
	Dim element1 As Object
	Dim item1 As Object
    Dim oOcc As Object
'    oAsmCD.OccurrencePatterns
	
    For Each oOcc In ThisApplication.ActiveDocument.ComponentDefinition.OccurrencePatterns
        Try
            	If Component.IsActive(oOcc.Name) = False Then
					For Each element1 In oOcc.OccurrencePatternElements
						For Each item1 In element1.Occurrences
							If Component.IsActive(item1.Name) = False Then
								oOcc.Delete
							End If
						Next
					Next
					
            	End If
        Catch ex As Exception
        End Try
    Next
	
End Sub



Function IsPatternWithUnsuppressedElements(oOcc As ComponentOccurrence) As Boolean
    ' Check if the occurrence is part of a pattern and has any unsuppressed elements
    If oOcc.IsPatternElement Then
        Dim oPattern As Object
        oPattern = oOcc.PatternElement.Parent
        Dim oPatternElem As Object
        
        ' Check each element in the pattern for suppression state
        For Each oPatternElem In oPattern.PatternElements
            If Not oPatternElem.Suppressed Then
                Return True ' Found an unsuppressed element, don't delete the pattern
            End If
        Next
    End If
    Return False ' All elements are suppressed or it’s not a pattern
End Function


Function GetAssyLevelItemToDelete(oOcc As Object) As Object
    If oOcc.IsPatternElement Then
        oOcc = oOcc.PatternElement.Parent
        oOcc = GetAssyLevelItemToDelete(oOcc)
    End If
    Return oOcc
End Function


	
	

