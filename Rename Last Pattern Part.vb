
Dim oDef As AssemblyComponentDefinition = ThisApplication.ActiveDocument.ComponentDefinition
Dim oPattern As OccurrencePattern
Dim x As Integer
Dim result As Integer

For Each oPattern In oDef.OccurrencePatterns

    If oPattern.Name = "Ceiling Flash Column 1" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            ' Rename the last occurrence in the pattern
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 1"
            End If
        End If
        result = x
    End If

    If oPattern.Name = "Ceiling Flash Column 2" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 2"
            End If
        End If
        result = x
    End If
	
	If oPattern.Name = "Ceiling Flash Column 3" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            ' Rename the last occurrence in the pattern
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 3"
            End If
        End If
        result = x
    End If

    If oPattern.Name = "Ceiling Flash Column 4" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 4"
            End If
        End If
        result = x
    End If
	
	If oPattern.Name = "Ceiling Flash Column 5" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            ' Rename the last occurrence in the pattern
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 5"
            End If
        End If
        result = x
    End If

    If oPattern.Name = "Ceiling Flash Column 6" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 6"
            End If
        End If
        result = x
    End If
	
	If oPattern.Name = "Ceiling Flash Column 7" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            ' Rename the last occurrence in the pattern
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 7"
            End If
        End If
        result = x
    End If

    If oPattern.Name = "Ceiling Flash Column 8" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 8"
            End If
        End If
        result = x
    End If
	
	If oPattern.Name = "Ceiling Flash Column 9" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            ' Rename the last occurrence in the pattern
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 9"
            End If
        End If
        result = x
    End If

    If oPattern.Name = "Ceiling Flash Column 10" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 10"
            End If
        End If
        result = x
    End If
	
	If oPattern.Name = "Ceiling Flash Column 11" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            ' Rename the last occurrence in the pattern
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 11"
            End If
        End If
        result = x
    End If

    If oPattern.Name = "Ceiling Flash Column 12" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 12"
            End If
        End If
        result = x
    End If
	
	If oPattern.Name = "Ceiling Flash Column 13" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            ' Rename the last occurrence in the pattern
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 13"
            End If
        End If
        result = x
    End If

    If oPattern.Name = "Ceiling Flash Column 14" Then
        x = oPattern.OccurrencePatternElements.Count
        If x > 0 Then
            Dim oLastElement As OccurrencePatternElement = oPattern.OccurrencePatternElements.Item(x)
            Dim occCount As Integer = oLastElement.Occurrences.Count
            If occCount > 0 Then
                oLastElement.Occurrences.Item(occCount).Name = "Last Ceiling 14"
            End If
        End If
        result = x
    End If
	
Next

