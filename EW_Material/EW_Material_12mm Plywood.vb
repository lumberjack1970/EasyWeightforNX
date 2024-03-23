' EasyWeight
' Journal desciption: Changes body color, layer and translucency, sets a density value, measures volume, calculates weight, and attributes it: EW_Body_Weight, EW_Material_Density and EW_Material.
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13
' V100 - Initial Release - November 2023
' V101 - Unit system support and Configuration Settings

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
    Dim theUI As UI
    Dim bodyWeight As Double

    '------------------------
    ' Configuration Settings:

    ' Material Name - Change this name to your own material
    Dim materialname As String = "12mm Plywood"

    ' Material Density - Kg/m3 or Pound/Cubic Foot - Change this value to your own specific density
    Dim density As Double = 440

    ' Unit System - "kg" for Kilograms or "lbm" for Pounds.
    Dim unitsystem As String = "kg"

    ' Body Settings:
    Dim bodycolor As Double = 111 ' Set the solid body color to ID: 111
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

            Dim markId1 As Session.UndoMarkId = Nothing
            markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Material Journal")

            If SelectObjects("Hey, select multiple somethings", mySelectedObjects) = Selection.Response.Ok Then
                For Each tempComp As Body In mySelectedObjects
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

                    DeleteAllAttributes(tempComp)
                    AddBodyAttribute(tempComp, "EW_Material", materialname)

                    Dim attributePropertiesBuilder1 As AttributePropertiesBuilder
                    attributePropertiesBuilder1 = theSession.AttributeManager.CreateAttributePropertiesBuilder(workPart, {tempComp}, AttributePropertiesBuilder.OperationType.None)
                    attributePropertiesBuilder1.Title = "EW_Material_Density"
                    attributePropertiesBuilder1.DataType = NXOpen.AttributePropertiesBaseBuilder.DataTypeOptions.Number

                    If unitsystem = "kg" Then
                        attributePropertiesBuilder1.Units = "KilogramPerCubicMeter"
                    Else
                        attributePropertiesBuilder1.Units = "PoundMassPerCubicFoot"
                    End If

                    attributePropertiesBuilder1.NumberValue = density

                    AddBodyAttribute(tempComp, "Component_created", String.Empty)

                    Dim nXObject1 As NXObject
                    nXObject1 = attributePropertiesBuilder1.Commit()
                    attributePropertiesBuilder1.Destroy()

                    ' Calculate volume using mass properties
                    Dim myMeasure As MeasureManager = workPart.MeasureManager()
                    Dim massUnits(4) As Unit
                    massUnits(1) = workPart.UnitCollection.GetBase("Volume")
                    Dim mb As MeasureBodies = myMeasure.NewMassProperties(massUnits, 0.99, {tempComp})

                    ' Create or update an attribute named 'SS_Body_Weight' and assign the weight value to it
                    Dim attributePropertiesBuilderForWeight As AttributePropertiesBuilder
                    attributePropertiesBuilderForWeight = theSession.AttributeManager.CreateAttributePropertiesBuilder(workPart, {tempComp}, AttributePropertiesBuilder.OperationType.Create)

                    If unitsystem = "kg" Then
                        mb.InformationUnit = MeasureBodies.AnalysisUnit.KilogramMilliMeter
                        ' Extract volume and multiply it by density to get weight
                        Dim bodyVolume As Double = mb.Volume
                        bodyWeight = bodyVolume / 1000000000.0 * density
                        attributePropertiesBuilderForWeight.Units = "Kilogram"
                    Else
                        mb.InformationUnit = MeasureBodies.AnalysisUnit.PoundInch
                        Dim bodyVolume As Double = mb.Volume
                        bodyWeight = bodyVolume / 1728 * density
                        attributePropertiesBuilderForWeight.Units = "Lbm"
                    End If

                    attributePropertiesBuilderForWeight.Title = "EW_Body_Weight"
                    attributePropertiesBuilderForWeight.DataType = AttributePropertiesBaseBuilder.DataTypeOptions.Number
                    attributePropertiesBuilderForWeight.NumberValue = bodyWeight
                    Dim attributeObjectForWeight As NXObject = attributePropertiesBuilderForWeight.Commit()
                    attributePropertiesBuilderForWeight.Destroy()
                Next
            End If
        Catch ex As Exception
            Console.WriteLine("Houston, we have a situation... an error occurred: " & ex.Message)
        End Try
    End Sub

    Function SelectObjects(prompt As String,
                           ByRef dispObj As List(Of DisplayableObject)) As Selection.Response
        Dim selObj As NXObject()
        Dim title As String = ("Select solid bodies - " & materialname)
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

    Sub DeleteAllAttributes(ByVal theObject As NXObject)
        Dim attributeInfo() As NXObject.AttributeInformation = theObject.GetUserAttributes()

        For Each temp As NXObject.AttributeInformation In attributeInfo
            theObject.DeleteUserAttributes(temp.Type, Update.Option.Now)
        Next
    End Sub

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
    End Function
End Module
