Sub Main()
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

