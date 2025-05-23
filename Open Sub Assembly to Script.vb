Dim Width As Integer, Length As Integer
Width = Parameter("TankWidth") + 1
Length = Parameter("TankLength") + 1

Parameter("TankBaseAssembly:1", "TankWidth") = Width
Parameter("TankBaseAssembly:1", "TankLength") = Length


Dim iLogicAuto As Object
Dim oRule As Object
iLogicAuto = iLogicVb.Automation
Dim RuleName1 As String = "Delete All"
Dim RuleName2 As String = "Execute Automation"

' Get the parent assembly document
Dim parentAsm As AssemblyDocument
parentAsm = ThisApplication.ActiveDocument

' Find the child assembly occurrence by name
Dim childAsmOccurrence As ComponentOccurrence
childAsmOccurrence = parentAsm.ComponentDefinition.Occurrences.ItemByName("TankBaseAssembly:1")

' Activate the child assembly
Dim childAsm As AssemblyDocument
childAsm = childAsmOccurrence.Definition.Document

 ThisApplication.Documents.Open(childAsm.FullDocumentName)
' Run the iLogic rule in the context of the child assembly
childAsm.Activate()


oRule = iLogicAuto.GetRule(childAsm, RuleName1)
iLogicAuto.RunRuleDirect(oRule)

oRule = iLogicAuto.GetRule(childAsm, RuleName2)
iLogicAuto.RunRuleDirect(oRule)


childAsm.Close(True)
