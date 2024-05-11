' EasyWeight
' Journal desciption: Changes body color, layer and translucency, and sets a NX's built in material.
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13
' V100 - Initial Release - November 2023
' V101 - Code cleanup to focus mainly on NX's built in material and added Configuration Settings

Imports System
Imports NXOpen
Imports System.Collections.Generic
Imports NXOpen.UF

Module NXJournal
    Dim theSession As Session
    Dim theUFSession As UFSession
    Dim workPart As Part
    Dim displayPart As Part
    Dim mySelectedObjects As New List(Of DisplayableObject)
	Dim physicalMaterial1 As PhysicalMaterial = Nothing
	Dim theUI As UI


	'------------------------
	' Configuration Settings:

	' Material Name - This name has to match with your created material in your library
	Dim materialName As String = "Fabric" 

	' Material Library Path
	Dim materialLibraryPath As String = "C:\Your Folder\To your material library\physicalmateriallibrary_custom.xml"

	' Body Settings:
	Dim bodycolor As Double = 45 ' Set the solid body color to ID: 45
	Dim bodylayer As Double = 1 ' Set the solid body to layer 1
	Dim bodytranslucency As Double = 0 ' Set the solid body transparency to 0
	'------------------------


	Sub Main(ByVal args() As String)

		Try
			theSession = Session.GetSession()
			theUFSession = UFSession.GetUFSession()
			workPart = theSession.Parts.Work
			displayPart = theSession.Parts.Display
			theUI = UI.GetUI()
			Dim lw As ListingWindow = theSession.ListingWindow
			lw.Open()

			Dim markId1 As Session.UndoMarkId = Nothing
			markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Material Journal")

			' Initialize the builders
			Dim physicalMaterialListBuilder1 As NXOpen.PhysMat.PhysicalMaterialListBuilder = Nothing
			Dim physicalMaterialAssignBuilder1 As NXOpen.PhysMat.PhysicalMaterialAssignBuilder = Nothing
			physicalMaterialListBuilder1 = workPart.MaterialManager.PhysicalMaterials.CreateListBlockBuilder()
			physicalMaterialAssignBuilder1 = workPart.MaterialManager.PhysicalMaterials.CreateMaterialAssignBuilder()
			Dim materialLibraryLoaded As Boolean = False
			Try
				' Check if the material is already loaded
				Dim loadedMaterial As NXOpen.Material = workPart.MaterialManager.PhysicalMaterials.GetLoadedLibraryMaterial(materialLibraryPath, materialName)
				If loadedMaterial Is Nothing Then
					' Material is not loaded, load it from the library
					physicalMaterial1 = workPart.MaterialManager.PhysicalMaterials.LoadFromMatmlLibrary(materialLibraryPath, materialName)
					materialLibraryLoaded = True
					'lw.WriteLine("Material library loaded.")
				Else
					physicalMaterial1 = loadedMaterial
					materialLibraryLoaded = True
					'lw.WriteLine("Material already loaded.")
				End If
			Catch ex As Exception
				lw.WriteLine("Failed to check/load material library: " & ex.Message)
			End Try

			If SelectObjects("Hey, select multiple somethings", mySelectedObjects) = Selection.Response.Ok Then
				For Each tempComp As Body In mySelectedObjects
					Try
						Dim pmaterialname as String = ("PhysicalMaterial[" & materialName & "]")
						Dim physicalMaterial1 As NXOpen.PhysicalMaterial = CType(workPart.MaterialManager.PhysicalMaterials.FindObject(pmaterialname), NXOpen.PhysicalMaterial)
						If physicalMaterial1 Is Nothing Then
							lw.WriteLine("Error: Material " & materialName & " not found.")
							Return
						End If

						Dim displayModification1 As DisplayModification
						displayModification1 = theSession.DisplayManager.NewDisplayModification()

						With displayModification1
							.ApplyToAllFaces = False
							.ApplyToOwningParts = True
							.NewColor = bodycolor
							.NewLayer = bodylayer 
							.NewTranslucency = bodytranslucency 
							.Apply({tempComp})
						End With

						displayModification1.Dispose()

						AddBodyAttribute(tempComp, "Component_created", String.Empty)

						If physicalMaterial1 IsNot Nothing Then
							physicalMaterial1.AssignObjects(New NXOpen.NXObject() {tempComp})
							'lw.WriteLine("Material " & materialName & " successfully assigned to body: " & tempComp.JournalIdentifier)
						Else
							lw.WriteLine("Error: Material " & materialName & " not found in the material library.")
						End If
					Catch ex As Exception
						lw.WriteLine("Error processing body: " & tempComp.JournalIdentifier & " - " & ex.Message)
						lw.WriteLine("Exception occurred: " & ex.ToString())
					End Try
				Next
				theSession.UpdateManager.DoUpdate(markId1)
			Else
				lw.WriteLine("No objects were selected.")
			End If

			If physicalMaterialAssignBuilder1 IsNot Nothing Then
				physicalMaterialAssignBuilder1.Destroy()
				physicalMaterialAssignBuilder1 = Nothing
			End If

			If physicalMaterialListBuilder1 IsNot Nothing Then
				physicalMaterialListBuilder1.Destroy()
				physicalMaterialListBuilder1 = Nothing
			End If
		Catch ex As Exception
			Console.WriteLine("Houston, we have a situation... an error occurred: " & ex.Message)
		Finally
			If physicalMaterial1 IsNot Nothing Then
				physicalMaterial1 = Nothing
			End If
        End Try
    End Sub

    Function SelectObjects(prompt As String,
                           ByRef dispObj As List(Of DisplayableObject)) As Selection.Response
        Dim selObj As NXObject()
		Dim title As String = ("Select solid bodies - " & materialName)
		Dim includeFeatures As Boolean = False
        Dim keepHighlighted As Boolean = False
        Dim selAction As Selection.SelectionAction = Selection.SelectionAction.ClearAndEnableSpecific
        Dim scope As Selection.SelectionScope = Selection.SelectionScope.WorkPart
        Dim selectionMask_array(0) As Selection.MaskTriple

        With selectionMask_array(0)
            .Type = UFConstants.UF_solid_type
            .SolidBodySubtype = UFConstants.UF_UI_SEL_FEATURE_SOLID_BODY
        End With

        Dim resp As Selection.Response = theUI.SelectionManager.SelectObjects(prompt,
            title, scope, selAction,
            includeFeatures, keepHighlighted, selectionMask_array,
            selObj)

        If resp = Selection.Response.ObjectSelected Or
                resp = Selection.Response.ObjectSelectedByName Or
                resp = Selection.Response.Ok Then
            For Each item As NXObject In selObj
                dispObj.Add(CType(item, DisplayableObject))
            Next
            Return Selection.Response.Ok
        Else
            Return Selection.Response.Cancel
        End If
    End Function

	Sub AddBodyAttribute(ByVal theBody As Body, ByVal attTitle As String, ByVal attValue As String)
		Dim attributePropertiesBuilder1 As AttributePropertiesBuilder
		attributePropertiesBuilder1 = theSession.AttributeManager.CreateAttributePropertiesBuilder(workPart, {theBody}, AttributePropertiesBuilder.OperationType.None)

		attributePropertiesBuilder1.IsArray = False
		attributePropertiesBuilder1.DataType = AttributePropertiesBaseBuilder.DataTypeOptions.String

		attributePropertiesBuilder1.Title = attTitle
		attributePropertiesBuilder1.StringValue = attValue

		Dim nXObject1 As NXObject
		nXObject1 = attributePropertiesBuilder1.Commit()

		attributePropertiesBuilder1.Destroy()
	End Sub

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
    End Function
End Module
