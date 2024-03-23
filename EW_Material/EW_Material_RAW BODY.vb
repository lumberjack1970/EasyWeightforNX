' EasyWeight
' Journal desciption: By selecting the original body and the raw body, this calculates the weight difference and adds a new attribute: Raw_Body_Delta_Weight. It also moves the raw body to a predefined layer and makes it transparent.
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13
' V100 - Initial Release - November 2023
' V101 - Unit system support and added Configuration Settings

Imports System
Imports NXOpen
Imports System.Collections.Generic
Imports NXOpen.UF

Module NXJournal
    Dim theSession As Session = Session.GetSession()
    Dim theUFSession As UFSession
    Dim workPart As Part
    Dim displayPart As Part
    Dim mySelectedObjects As New List(Of DisplayableObject)
    Dim theUI As UI
	Dim lw As ListingWindow = theSession.ListingWindow
	Dim scrib_weight As Double

	'------------------------
	' Configuration Settings:

	' Body Settings:
	Dim bodyLayer As Integer = 170 ' Set the solid body to layer 1
	Dim bodyTranslucency As Integer = 70 ' Set the solid body transparency to 70
	'------------------------

	Sub Main(ByVal args() As String)

		Dim theSession As Session = Session.GetSession()
		Dim markId1 As NXOpen.Session.UndoMarkId
		markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Raw Body Journal")

		lw.Open()
		Try
			theSession = Session.GetSession()
			theUFSession = UFSession.GetUFSession()
			workPart = theSession.Parts.Work
			displayPart = theSession.Parts.Display
			theUI = UI.GetUI()

			' Variables to store attributes from the first selected body (original body)
			Dim origmat_density As Double = 0.0
			Dim orig_weight As Double = 0.0

			' First selection for the original body
			Try
				If SelectObjects("Select the Original body", mySelectedObjects) = Selection.Response.Ok Then
					For Each tempComp As Body In mySelectedObjects
						Dim attributes = tempComp.GetUserAttributes()
						For Each attr1 As NXObject.AttributeInformation In attributes
							If attr1.Title = "EW_Material_Density" Then
								origmat_density = CDbl(attr1.StringValue)
							ElseIf attr1.Title = "EW_Body_Weight" Then
								orig_weight = CDbl(attr1.StringValue)
							End If
						Next
					Next
				End If
				If origmat_density = 0.0 Or orig_weight = 0.0 Then
					'Throw New Exception("Failed to decipher the weight value for this body! Please apply the appropriate Material Journal before proceeding")
				End If

				' Clear previous selections
				mySelectedObjects.Clear()

				' Second selection for the Raw body
				If SelectObjects("Select the Raw body", mySelectedObjects) = Selection.Response.Ok Then
					For Each tempComp As Body In mySelectedObjects
						' Measure Raw body volume
						Dim myMeasure As MeasureManager = workPart.MeasureManager()
						Dim massUnits(4) As Unit
						massUnits(1) = workPart.UnitCollection.GetBase("Volume")

						Dim mb As MeasureBodies = myMeasure.NewMassProperties(massUnits, 0.99, {tempComp})

						Dim bodyVolume As Double = mb.Volume
						If bodyVolume = 0.0 Then
							Throw New Exception("Invalid body volume.")
						End If

						If Double.IsNaN(scrib_weight) Or Double.IsInfinity(scrib_weight) Then
							Throw New Exception("Invalid Raw weight calculated.")
						End If

						' Add new attribute
						Dim attributePropertiesBuilderForWeight As AttributePropertiesBuilder
						attributePropertiesBuilderForWeight = theSession.AttributeManager.CreateAttributePropertiesBuilder(workPart, {tempComp}, AttributePropertiesBuilder.OperationType.Create)
						attributePropertiesBuilderForWeight.Title = "Raw_Body_Delta_Weight"
						attributePropertiesBuilderForWeight.DataType = AttributePropertiesBaseBuilder.DataTypeOptions.Number

						' Now, perform the unit check on the currentPart
						If workPart.PartUnits = BasePart.Units.Inches Then

							mb.InformationUnit = MeasureBodies.AnalysisUnit.PoundInch
							attributePropertiesBuilderForWeight.Units = "Lbm"
							scrib_weight = origmat_density * (bodyVolume / 1728) - orig_weight
						Else

							mb.InformationUnit = MeasureBodies.AnalysisUnit.KilogramMilliMeter
							attributePropertiesBuilderForWeight.Units = "Kilogram"
							scrib_weight = origmat_density * (bodyVolume / 1000000000.0) - orig_weight
						End If

						attributePropertiesBuilderForWeight.NumberValue = scrib_weight
						AddBodyAttribute(tempComp, "Component_created", String.Empty)
						Dim attributeObjectForWeight As NXObject = attributePropertiesBuilderForWeight.Commit()
						attributePropertiesBuilderForWeight.Destroy()

						' Delete old attribute
						DeleteUserAttribute(tempComp, "EW_Body_Weight")

						' Move body to Layer 130
						Dim displayModification1 As DisplayModification
						displayModification1 = theSession.DisplayManager.NewDisplayModification()
						With displayModification1
							.ApplyToAllFaces = False
							.ApplyToOwningParts = True
							.NewTranslucency = bodyTranslucency
							.NewLayer = bodyLayer
							.Apply({tempComp})
						End With
						displayModification1.Dispose()
					Next
				End If
			Catch ex As Exception
				lw.WriteLine("An error occurred: " & ex.Message)
			Finally
				lw.Close()
			End Try

		Catch ex As Exception
			Console.WriteLine("An error occurred: " & ex.Message)
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

    Sub DeleteUserAttribute(ByVal theObject As NXObject, ByVal attributeName As String)
		Dim attributeInfo() As NXObject.AttributeInformation = CType(theObject, Body).GetUserAttributes()
		
		For Each temp As NXObject.AttributeInformation In attributeInfo
			If temp.Title = attributeName Then
				theObject.DeleteUserAttribute(temp.Type, temp.Title, False, Update.Option.Now)
				Exit For
			End If
		Next
	End Sub

	Public Function GetUnloadOption(ByVal dummy As String) As Integer
        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
    End Function
End Module