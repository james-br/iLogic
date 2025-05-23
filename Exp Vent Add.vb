Sub Main()
	iLogicVb.RunRule("Rule 41: Deflagration Calculator")
	Dim tempcount As Integer = 0

	Do While True
        ' Your code here
                
        ' Allow other events to process
        Dim doc As Document
	    doc = ThisApplication.ActiveDocument
	    Dim VentQTY As Integer
	    VentQTY = 0
	            
	    ' Pick an occurrence in the assembly
	    Dim enti As Object
	    On Error Resume Next
	    enti = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Select a ceiling panel. Press Esc to Exit.")
	    On Error GoTo 0

	    If enti Is Nothing Then
	        MsgBox ("No entity selected or operation canceled.")
	        Exit Sub
	    End If
	
		'==================================================
		'Checks if the right item was selected
		If InStr(enti.Name, "-STANDARD") > 0  Then
	      	
		Else
	    	MsgBox("Please select a Ceiling Panel")
			Exit Sub
		End If
	
		'==================================================
		'Finds the element inside the Pattern
	    ' Checks if the selected entity is a proxy and get the actual occurrence
	    Dim occurrence As ComponentOccurrence
	    If TypeOf enti Is ComponentOccurrenceProxy Then
	        occurrence = enti.ContainingOccurrence
	    ElseIf TypeOf enti Is ComponentOccurrence Then
	        occurrence = enti
	    Else
	        MsgBox ("Selected entity is not a valid occurrence.")
	        Exit Sub
	    End If
	
	    ' Get the name of the selected entity
	    Dim entityName As String
		Dim placeholder As String
	    entityName = occurrence.Name
		
		'EntityName is the component name
		'==================================================
		Dim patternComponent1 As New Collection
		Dim patternComponent2 As New Collection
		patternComponent1 = findAssembly(entityName)
		
		
		If Not patternComponent1 Is Nothing Then
			
			If tempcount >= CInt(DefragCalculator) Then
		
				Dim response As Integer
				response = MsgBox("You have created sufficient vents for this oven: " & DefragCalculator & ". Do you want to add more?", vbYesNo)
				
				' Respond based on the user's choice
				If response = vbYes Then
				    GoTo ER
				Else
				    Exit Sub
				End If
			End If
ER:			
			placeholder = RemoveVent(patternComponent1)
			tempcount = ImportVent(placeholder,tempcount)
		End If
		
    Loop
	
	
End Sub

Function RemoveVent(patternComponentC As Collection)As String
	Dim pattern As Object
    Dim pattern1 As Object
    Dim pattern2 As Object
    Dim occ As ComponentOccurrence
    Dim patternFeatureName As String
	Dim solution As Collection
	
    For Each pattern In ThisApplication.ActiveDocument.ComponentDefinition.OccurrencePatterns
		If pattern.name = patternComponentC(3) Then
			Dim previousElement As Object
			Dim lastElement As String
			Dim replaceElement As String
			Dim elementCount As Integer
			Dim pattern3 As Object
			Dim pattern4 As Object
			
			elementCount = pattern.OccurrencePatternElements.Count
			lastElement = "Element:" & elementCount ' Construct the name of the last element based on the count
			secondLastElement = "Element:" & elementCount - 1 ' Construct the name of the last element based on the count
			
			
	        For Each pattern1 In pattern.OccurrencePatternElements
				If pattern1.name = patternComponentC(2) Then
					If pattern1.name ="Element:1" Then
						MsgBox(pattern1.name)
						Exit Function
					End If
					If pattern1.name = lastElement Then
						MsgBox(pattern1.name)
						MsgBox("last element")
						Exit Function
					End If
					If pattern1.name = secondLastElement Then
						MsgBox(pattern1.name)
						MsgBox("Second last element")
						Exit Function
					End If
										
					
					 For i = 1 To elementCount
			            If pattern.OccurrencePatternElements(i).name = patternComponentC(2) Then							 
			                If i < elementCount Then
								Dim secondItemS As Collection
								Dim holdingItemS As Collection
			                    pattern3 = pattern.OccurrencePatternElements(i + 1).name
								secondItemS = secondItem(patternComponentC, pattern3)
								pattern4 = pattern.OccurrencePatternElements(i - 1).name
								holdingItemS = secondItem(patternComponentC, pattern4)
						
								
								Component.IsActive(patternComponentC(1)) = False	
								Component.IsActive(secondItemS(1)) = False	
								RemoveVent = holdingItemS(1)
								Return RemoveVent 
			                End If
			            End If
			        Next i
				End If
			Next pattern1
		End If
    Next pattern
		
End Function




Function secondItem(firstItem As Collection, itemName As String) As Collection
    'Find the second item feature name
    Dim solution As New Collection
    Dim pattern As Object
    Dim pattern1 As Object
    Dim pattern2 As Object
	Dim pattern3 As Object
    Dim occ As ComponentOccurrence
    Dim patternFeatureName As String
    
    For Each pattern In ThisApplication.ActiveDocument.ComponentDefinition.OccurrencePatterns
		If pattern.name = firstItem(3) Then			
	        For Each pattern1 In pattern.OccurrencePatternElements
				If pattern1.name = itemName Then					
		            For Each pattern2 In pattern1.Occurrences						
						solution.Add(pattern2.Name)
						solution.Add(pattern1.Name)
						solution.Add(pattern.Name)
						secondItem = solution
		            	Exit Function
		            Next pattern2
				End If
        	Next pattern1
		End If
    Next pattern
	
    MsgBox ("Component " & itemName & " is not part of any pattern feature.")
    secondItem = Nothing
End Function


Function findAssembly(itemName As String) As Collection
    'Find the pattern feature name
    Dim solution As New Collection
    Dim pattern As Object
    Dim pattern1 As Object
    Dim pattern2 As Object
    Dim occ As ComponentOccurrence
    Dim patternFeatureName As String
    
    For Each pattern In ThisApplication.ActiveDocument.ComponentDefinition.OccurrencePatterns
        For Each pattern1 In pattern.OccurrencePatternElements
            For Each pattern2 In pattern1.Occurrences
                If pattern2.Name = itemName Then
					If pattern1.Name = "Element:1"  Or pattern1.Index = pattern.OccurrencePatternElements.Count Or pattern1.Index = pattern.OccurrencePatternElements.Count -1 Or pattern.Name = "Last Ceiling Row" Or pattern.Name = "First Ceiling Row" or InStr(pattern.Name, "Middle Ceiling Row") = 0 Then
						MsgBox("Please select an inner ceiling panel. ")
						findAssembly = Nothing
						Exit Function
					End If
                    solution.Add(pattern2.Name)
                    solution.Add(pattern1.Name)
                    solution.Add(pattern.Name)
					
                    findAssembly = solution
                    Exit Function
                End If
            Next pattern2
        Next pattern1
    Next pattern

    MsgBox ("Component " & itemName & " is not part of any pattern feature.")
    findAssembly = Nothing
End Function



Function ImportVent(itemName As String, tempcount1 As Integer) As Integer
	' Define the component names (assuming no spaces in names)
    Dim newComponentName As String
    newComponentName = "OPX4312060-TEST" ' Include revision if necessary
	
    ' Define the file path of the component to be imported
    Dim filePath As String
    filePath = "C:\Vault Works\Designs\Panels & Parts\OTYPE\" & newComponentName & ".iam" ' Build the path from component name
	
	' Check if the file exists before importing
    If Dir(filePath) = "" Then
        MsgBox ("Error: The file " & filePath & " does not exist.")
        Exit Function
    End If
	
    ' Get the assembly document
    Dim doc As AssemblyDocument
    doc = ThisApplication.ActiveDocument

    ' Get the component occurrences in the assembly
    Dim occurrences As ComponentOccurrences
    occurrences = doc.ComponentDefinition.Occurrences

    
    ' Import the component if it does not exist
    Dim transMatrix As Matrix
    transMatrix = ThisApplication.TransientGeometry.CreateMatrix
    transMatrix.SetToIdentity

    Dim newOccurrence As ComponentOccurrence
    newOccurrence = occurrences.Add(filePath, transMatrix)
    Dim insertValues As Integer = GetNumberAfterColon(itemName)
    
    If Not itemName Is Nothing Then

        Constraints.AddByiMates ("Insert:" & insertValues + 1 , newOccurrence.Name, "OP-OPX1", itemName, "OP-OPX4")
        Constraints.AddByiMates("Insert:" & insertValues + 2, newOccurrence.Name, "OP-OPX2", itemName, "OP-OPX3")
		tempcount1 = tempcount1 + 1
		doc.Update
		Return tempcount1
    Else
        MsgBox ("The component " & existingComponentName & " was not found.")
  
    End If
	
End Function


Function GetNumberAfterColon(InputString As String) As String
	
    Dim parts() As String
    If TypeName(InputString) = "String" Then
		GetCurrentDateTime = Format(Now, "yyMMddss")
		Return GetCurrentDateTime
	End If
		
	parts = Split(InputString, ":")
    
    If UBound(parts) >= 1 Then
        GetNumberAfterColon = Trim(parts(1))
    Else
        GetNumberAfterColon = ""
    End If
End Function


