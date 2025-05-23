
Dim oAsmDoc As AssemblyDocument
oAsmDoc = ThisApplication.ActiveDocument

Dim oOcc As ComponentOccurrence
Dim oSubAsm As PartDocument
Dim subAsmPath As String
Dim newSubAsmPath As String


For Each oOcc In oAsmDoc.ComponentDefinition.Occurrences
    If oOcc.Name = "OvenReferenceExtrusion:1" Then 'Or _
'	   oOcc.Name = "Vertical Flashing Trim Extrusion" Or _
'	   oOcc.Name = "Horizontal Flashing Trim Extrusion L to R" Or _
'	   oOcc.Name = "Horizontal Flashing Trim Extrusion F to B" Then
		
        oSubAsm = oOcc.Definition.Document
        subAsmPath = oSubAsm.FullFileName
        'MsgBox("Original Path: " & subAsmPath)

        Dim baseName As String
        baseName = Parameter("FileTitle") & "-" &  oOcc.Name

        If baseName = "" Then
            MsgBox("Parameter BaseName is empty!", , "Error")
            Exit Sub
        End If

        Dim invalidChars As String = "\/:*?""<>|"
        Dim badChar As String
        For Each badChar In invalidChars
            If baseName.Contains(badChar) Then
                MsgBox("Invalid character '" & badChar & "' in BaseName.", , "Invalid File Name")
                Exit Sub
            End If
        Next

        newSubAsmPath = "C:\Vault Works\Designs\Units\Ovens\Oven Panels & Parts\" & baseName & ".ipt"

        If Not System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(newSubAsmPath)) Then
            MsgBox("Folder does not exist: " & System.IO.Path.GetDirectoryName(newSubAsmPath), , "Folder Missing")
            Exit Sub
        End If

        MsgBox("Saving to: " & newSubAsmPath)

        ' Save and Replace
        ' Close any edited versions first if needed
		Try
			System.IO.File.Copy(subAsmPath, newSubAsmPath, False)
			MsgBox("Copied")
		Catch ex As Exception
			MsgBox("Error: " & ex.Message)
		End Try
		Try
	        oOcc.Replace(newSubAsmPath, True)
		Catch ex As Exception
			MsgBox("Error: " & ex.Message)
		End Try
		Exit Sub
    End If
Next

MsgBox("Update Complete. Oven Trim Extrusions were updated.", , "iLogic")
iLogicVb.UpdateWhenDone = True
