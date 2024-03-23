' EasyWeight
' Journal desciption: By selecting the original body and the raw body, this calculates the weight difference and adds a new attribute: Raw_Body_Delta_Weight. It also moves the raw body to a predefined layer and makes it transparent.
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13
' V100 - Initial Release - November 2023
' V101 - Code cleanup to focus mainly on NX's built in material, Unit system support and added Configuration Settings

Imports System
Imports NXOpen
Imports System.Collections.Generic
Imports NXOpen.UF

Module NXJournal
    Dim theSession As Session = Session.GetSession()
    Dim workPart As Part
    Dim displayPart As Part
    Dim mySelectedObjects As New List(Of DisplayableObject)
    Dim physicalMaterial1 As PhysicalMaterial = Nothing
    Dim theUI As UI
    Dim lw As ListingWindow = theSession.ListingWindow

  
	'------------------------
	' Configuration Settings:

	' Material Name - This name has to match with your created material in your library
    Dim materialName As String = "NullMaterial"
	
	' Material Library Path
    Dim materialLibraryPath As String = "C:\Your Folder\To your material library\physicalmateriallibrary_custom.xml"
    
	' Body Settings:
	Dim bodyLayer As Integer = 130 ' Set the solid body to layer 1
    Dim bodyTranslucency As Integer = 70 ' Set the solid body transparency to 70
	'------------------------


    Sub Main(ByVal args() As String)

        Dim theSession As Session = Session.GetSession()
        Dim markId1 As NXOpen.Session.UndoMarkId
        markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Raw Body Journal")

        lw.Open()
        Try
            theSession = Session.GetSession()
            workPart = theSession.Parts.Work
            displayPart = theSession.Parts.Display
            theUI = UI.GetUI()

            ' Variables to store weight from the first and second selected body
            Dim orig_weight As Double = 0.0
            Dim Raw_weight As Double = 0.0

            ' First selection for the original body
            If SelectObjects("Select the Original body", mySelectedObjects) = Selection.Response.Ok Then
                For Each tempComp As Body In mySelectedObjects
                    orig_weight = GetBodyWeight(tempComp)
                Next
            End If

            mySelectedObjects.Clear()

            ' Second selection for the Raw body
            If SelectObjects("Select the Raw body", mySelectedObjects) = Selection.Response.Ok Then
                For Each tempComp As Body In mySelectedObjects
                    Raw_weight = GetBodyWeight(tempComp)
                    ' Calculate the weight difference
                    Dim weight_difference As Double = Raw_weight - orig_weight

                    ' Add new attribute for weight difference
                    Dim attributePropertiesBuilderForWeight As AttributePropertiesBuilder
                    attributePropertiesBuilderForWeight = theSession.AttributeManager.CreateAttributePropertiesBuilder(workPart, {tempComp}, AttributePropertiesBuilder.OperationType.Create)
                    attributePropertiesBuilderForWeight.Title = "Raw_Body_Delta_Weight"
                    attributePropertiesBuilderForWeight.DataType = AttributePropertiesBaseBuilder.DataTypeOptions.Number
                    attributePropertiesBuilderForWeight.NumberValue = weight_difference

                    ' Now, perform the unit check on the currentPart
                    If workPart.PartUnits = BasePart.Units.Inches Then
                        attributePropertiesBuilderForWeight.Units = "Lbm"
                    Else
                        attributePropertiesBuilderForWeight.Units = "Kilogram"
                    End If

                    Dim attributeObjectForWeight As NXObject = attributePropertiesBuilderForWeight.Commit()
                    attributePropertiesBuilderForWeight.Destroy()

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
                        'lw.WriteLine("Failed to check/load material library: " & ex.Message)
                    End Try

                    ' Apply display and material changes
                    Try
                        Dim pmaterialname As String = ("PhysicalMaterial[" & materialName & "]")
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
                            .NewLayer = bodyLayer
                            .NewTranslucency = bodyTranslucency
                            .Apply({tempComp})
                        End With

                        displayModification1.Dispose()

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

                    If physicalMaterialAssignBuilder1 IsNot Nothing Then
                        physicalMaterialAssignBuilder1.Destroy()
                        physicalMaterialAssignBuilder1 = Nothing
                    End If

                    If physicalMaterialListBuilder1 IsNot Nothing Then
                        physicalMaterialListBuilder1.Destroy()
                        physicalMaterialListBuilder1 = Nothing
                    End If
                Next
            End If
        Catch ex As Exception
            lw.WriteLine("An error occurred: " & ex.Message)
        Finally
            lw.Close()
        End Try
    End Sub


    Function SelectObjects(prompt As String,
                           ByRef dispObj As List(Of DisplayableObject)) As Selection.Response
        Dim selObj As NXObject() = Nothing
        Dim title As String = prompt
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

    Function GetBodyWeight(body As Body) As Double
        Dim weight As Double = 0.0
        Dim myMeasure As MeasureManager = theSession.Parts.Work.MeasureManager
        Dim theBodies(0) As Body
        theBodies(0) = body

        Dim massUnits(4) As NXOpen.Unit
        massUnits(2) = theSession.Parts.Work.UnitCollection.GetBase("Mass")

        Try
            Dim mb As MeasureBodies = myMeasure.NewMassProperties(massUnits, 0.99, theBodies)
            weight = mb.Mass
            mb.Dispose()
        Catch ex As Exception
            lw.WriteLine("Error measuring body weight: " & ex.Message)
        End Try
        Return weight
    End Function

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
    End Function
End Module
