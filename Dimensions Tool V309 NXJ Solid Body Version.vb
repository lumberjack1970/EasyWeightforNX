' Written by Tamas Woller - September 2024, V309
' Journal desciption: Iterates through and calculates the dimensions of each selected solid bodies - Length, Width And Material Thickness and attributes them.
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212

' ChangeLog:
' V100 - Initial Release - December 2023
' V101 - Improved the handling of NXOpen expressions
' V103 - Part-Level Unit Recognition, Measurement Precision Configuration, Nearest Half Rounding for Millimeters, Trim Trailing Zeros, GUI-Based Modification Control, Material Thickness Adjustment and added Configuration Settings
' V105 - Added "Maybe" to GUI-Based Modification Control
' V107 - Subassemblies (Components with children) skip, Part name output instead Component, added notes and minor changes in Output Window
' V109 - Expressions delete moved from the end to right after bounding box disposal
' V309 - Modified version to work with solid bodies instead of components. 

Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Utilities
Imports System.Drawing
Imports System.Windows.Forms

Module NXJournal
	Dim theSession As Session = Session.GetSession()
	Dim lw As ListingWindow = theSession.ListingWindow
	Dim theUFSession As UFSession = UFSession.GetUFSession()
	Dim theUI As UI = UI.GetUI()
	Dim workPart As Part = theSession.Parts.Work
	Dim displayPart As Part
	Dim mySelectedObjects As New List(Of DisplayableObject)
	Dim winFormLocationX As Integer = 317
	Dim winFormLocationY As Integer = 317
	Dim formattedWidth As Double
	Dim formattedDepth As Double
	Dim formattedHeight As Double
	Dim unitString As String
	Dim modifications As Boolean


	'------------------------
	' Configuration Settings:
	' Adjust these values accordingly, as the numbers represent the identified material thicknesses. The example below shows the most common thicknesses, with the first seven in inches and the following seven in millimeters. Note, the unit check occurs at the part level, allowing for different units (inch and millimeter) within the same assembly. The code automatically adjusts based on each part's unit.
	Dim validThicknesses As Double() = {0.141, 0.203, 0.25, 0.453, 0.469, 0.709, 0.827, 6, 9, 12, 13, 15, 18, 19}

	'Defines the precision for formatting measurements. "F0" means no decimal places, "F4" formats a number to 4 decimal places (e.g., 15.191832181891 to 15.1918). Values for NX makes sense between F0 to F13.
	Dim decimalFormat As String = "F1"

	'Controls whether to round measurements to the nearest half. Applicable only for millimeters and the decimal places will be forced to 1 by the code. When True, rounds 12.4999787 to 12.5. When False, rounds to the nearest whole number as decimalformat set (e.g. with F1, 12.2535. to 12.0, or with F3, 12.2535 to 12.254). Looks neat, if you use this with trimzeros set to True.
	Dim roundToNearestHalfBLN As Boolean = True

	' Determines the state of GUI-based modifications. A value of "True" enables user input through the GUI for modifications. A value of "False" bypasses the GUI, allowing the program to run automatically with predefined settings, akin to a "Just Do It" mode - no questions asked. "Maybe" will prompts you at the start of the journal to decide whether to enable modifications. Note, while the code is running, the NX window will not respond. Before starting, set your Model to trimetric view and position your Information Window so it doesn't obstruct your model. You can access the window using Ctrl+Shift+S.
	Dim modificationsQST As String = "Maybe"

	' Controls whether predefined material thickness adjustments are applied. When set to True, material thickness values are automatically adjusted to match predefined standards (e.g., for Laminates, you might model a 12mm Laminate with an extra 1mm for a total of 13mm, but require the output to be 12mm, thus the 13mm is adjusted back to 12mm). You need to include the "13" in the validThicknesses, so the code can identify at first, then adjust as required. To configure predefined values, see code from line 540 - under 'Configuration Settings'. You can add or remove values as needed, but ensure to maintain the format. 
	' Set to False to maintain original material thickness measurements without automatic adjustments. 
	Dim materialthicknessadjust As Boolean = True

	' Controls the trimming of unnecessary trailing zeros in the numerical output. When set to True, trailing zeros after the decimal point are removed for a cleaner display. For example, "12.34000" becomes "12.34", and "15.00" becomes "15". When False, numbers are displayed as formatted according to the specified decimal precision without trimming zeros.
	Dim trimzeros As Boolean = True

	' Quick setting variables for attribute names
	Dim LengthAttribute As String = "DIM_LENGTH"
	Dim WidthAttribute As String = "DIM_WIDTH"
	Dim MaterialThicknessAttribute As String = "MAT_THICKNESS"
	'------------------------

	
    Sub Main(ByVal args() As String)
		Dim nXObject1 As NXOpen.NXObject
		Dim bbWidth, bbDepth, bbHeight, Length, Width, MaterialThickness As Double
		Dim lengthDisplay As String = ""
		Dim widthDisplay As String = ""
		Dim markId1 As NXOpen.Session.UndoMarkId
		markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Dimensions Tool")
		lw.Open()

		' Assume modificationsQST is a String that can have values "True", "False", or anything else for an undecided state
		If modificationsQST IsNot Nothing AndAlso (modificationsQST.Equals("True", StringComparison.OrdinalIgnoreCase) OrElse modificationsQST.Equals("False", StringComparison.OrdinalIgnoreCase)) Then
			' If modificationsQST is a clear "True" or "False", convert it to a Boolean
			modifications = Boolean.Parse(modificationsQST)
		Else
			' If modificationsQST is not clearly "True" or "False", prompt the user
			Dim userResponse As DialogResult = MessageBox.Show("Would you like to modify the dimensions for these Solid Bodies?", "Modifications", MessageBoxButtons.YesNoCancel)
			Select Case userResponse
				Case DialogResult.Yes
					modifications = True
				Case DialogResult.No
					modifications = False
				Case DialogResult.Cancel
					lw.WriteLine(" ")
					lw.WriteLine("The incantation falters, no spell can be cast. But the dark arts wait for none, perhaps next time...")
					Return
			End Select
		End If
		
		lw.WriteLine("------------------------------------------------------------")
		lw.WriteLine("Lord Voldemort's Dimensions Tool          Version: 3.09 NXJ")
		lw.WriteLine("Solid Body Edition ")
		lw.WriteLine(" ")

		' Extract the number of decimal places from the format string for a clearer message
		Dim decimalPlaces As Integer = GetDecimalPlaces(decimalFormat)

		lw.WriteLine("--------------------------------")
		lw.WriteLine("Configuration Settings Overview:")
		lw.WriteLine(" ")
		lw.WriteLine("Numerical Output Configuration:")
		lw.WriteLine(" - Decimal Precision:             " & decimalPlaces.ToString())
		lw.WriteLine(" - Rounding to Nearest Half:      " & If(roundToNearestHalfBLN, "Yes", "No"))
		lw.WriteLine(" - Trimming Trailing Zeros:       " & If(trimzeros, "Yes", "No"))
		lw.WriteLine(" ")
		lw.WriteLine("Modifications:")
		lw.WriteLine(" - GUI Enabled:                   " & If(modifications, "Yes", "No"))

		Dim validThicknessesStr As String = ""
		For Each thickness As Double In validThicknesses
			If validThicknessesStr <> "" Then
				validThicknessesStr &= ", "
			End If
			validThicknessesStr &= thickness.ToString()
		Next

		lw.WriteLine(" - Material Thickness adjustment: " & If(materialthicknessadjust, "Yes", "No"))
		lw.WriteLine(" - Valid material thicknesses:    ")
		lw.WriteLine("   " & validThicknessesStr)
		lw.WriteLine(" ")
		lw.WriteLine("Attribute Names Configuration:")
		lw.WriteLine(" - Length Attribute:              " & LengthAttribute)
		lw.WriteLine(" - Width Attribute:               " & WidthAttribute)
		lw.WriteLine(" - Material Thickness Attribute:  " & MaterialThicknessAttribute)
		lw.WriteLine(" ")
		
        Try
            theSession = Session.GetSession()
            theUFSession = UFSession.GetUFSession()
            workPart = theSession.Parts.Work
            displayPart = theSession.Parts.Display
            theUI = UI.GetUI()

			If SelectObjects("Hey, select multiple somethings", mySelectedObjects) = Selection.Response.Ok Then
				' Initialize MeasureManager and unit handling for the entire part
				Dim currentPart As Part = workPart  ' Assuming you're working with the active work part
				Dim myMeasure As MeasureManager = currentPart.MeasureManager()

				' Create a unit array for the mass properties measurement
				Dim wIconMOseT(0) As Unit
				wIconMOseT(0) = currentPart.UnitCollection.GetBase("Length")

				' Create a mass properties measurement for all solid bodies in the part
				Dim mb As MeasureBodies = myMeasure.NewMassProperties(wIconMOseT, 0.99, currentPart.Bodies.ToArray())

				' Check the units of the current part and apply the corresponding analysis unit
				If currentPart.PartUnits = BasePart.Units.Inches Then
					unitString = "in"
					lw.WriteLine(" - Measurement Unit System: Imperial (Inches)")
					lw.WriteLine("-------------------------------")
					lw.WriteLine(" ")
					mb.InformationUnit = MeasureBodies.AnalysisUnit.PoundInch
				Else
					unitString = "mm"
					lw.WriteLine(" - Measurement Unit System: Metric (Millimeters)")
					mb.InformationUnit = MeasureBodies.AnalysisUnit.KilogramMilliMeter
				End If

				For Each obj As DisplayableObject In mySelectedObjects
					Dim tempComp As Body = TryCast(obj, Body)
					If tempComp Is Nothing Then
						lw.WriteLine(" ")
						lw.WriteLine("The solid body before me is flawed, incomplete in its essence.")
						Continue For
					End If

					Try
				        ' Write out the name of the solid body
						If String.IsNullOrEmpty(tempComp.Name) Then
							lw.WriteLine(" ")
							lw.WriteLine("-----------------------------")
							lw.WriteLine("Processing Overview for an unnamed Solid Body:")
						Else
							lw.WriteLine(" ")
							lw.WriteLine("-----------------------------")
							lw.WriteLine("Processing Overview for Solid Body:")
							lw.WriteLine(" - " & tempComp.Name)
						End If
						
						lw.WriteLine(" ")
						lw.WriteLine("The ritual begins... summoning the dark forces...")
						
						' Calculate and display the center of mass
						Dim accValue(10) As Double
						accValue(0) = 0.999
						Dim massProps(46) As Double
						Dim stats(12) As Double
						theUFSession.Modl.AskMassProps3d(New Tag() {tempComp.Tag}, 1, 1, 4, 0.03, 1, accValue, massProps, stats)

						' Convert the center of mass coordinates to Double
						Dim com_x As Double = massProps(3)
						Dim com_y As Double = massProps(4)
						Dim com_z As Double = massProps(5)

						Dim toolingBoxBuilder1 As NXOpen.Features.ToolingBoxBuilder = workPart.Features.ToolingFeatureCollection.CreateToolingBoxBuilder(Nothing)
						toolingBoxBuilder1.Type = NXOpen.Features.ToolingBoxBuilder.Types.BoundedBlock
						toolingBoxBuilder1.ReferenceCsysType = NXOpen.Features.ToolingBoxBuilder.RefCsysType.SelectedCsys
						toolingBoxBuilder1.XValue.SetFormula("10")
						toolingBoxBuilder1.YValue.SetFormula("10")
						toolingBoxBuilder1.ZValue.SetFormula("10")
						toolingBoxBuilder1.OffsetPositiveX.SetFormula("0")
						toolingBoxBuilder1.OffsetNegativeX.SetFormula("0")
						toolingBoxBuilder1.OffsetPositiveY.SetFormula("0")
						toolingBoxBuilder1.OffsetNegativeY.SetFormula("0")
						toolingBoxBuilder1.OffsetPositiveZ.SetFormula("0")
						toolingBoxBuilder1.OffsetNegativeZ.SetFormula("0")
						toolingBoxBuilder1.RadialOffset.SetFormula("0")
						toolingBoxBuilder1.Clearance.SetFormula("0")
						toolingBoxBuilder1.CsysAssociative = True
						toolingBoxBuilder1.NonAlignedMinimumBox = True
						toolingBoxBuilder1.SingleOffset = False

						Dim selectionIntentRuleOptions1 As NXOpen.SelectionIntentRuleOptions = Nothing
						selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions()
						selectionIntentRuleOptions1.SetSelectedFromInactive(False)

						Dim selectedBody As NXOpen.Body = TryCast(NXOpen.Utilities.NXObjectManager.Get(tempComp.Tag), NXOpen.Body)
						If selectedBody Is Nothing Then
							lw.WriteLine("The tag you dare present does not match any corporeal form.")
							Return
						End If

						' Use the selectedBody for creating the dumb rule
						Dim bodyDumbRule1 As NXOpen.BodyDumbRule = workPart.ScRuleFactory.CreateRuleBodyDumb(New Body() {selectedBody}, True, selectionIntentRuleOptions1)
						selectionIntentRuleOptions1.Dispose()

						Dim scCollector1 As NXOpen.ScCollector = toolingBoxBuilder1.BoundedObject
						Dim rules1(0) As NXOpen.SelectionIntentRule
						rules1(0) = bodyDumbRule1
						scCollector1.ReplaceRules(rules1, False)

						' Use the selectedBody in SetSelectedOccurrences
						Dim selections1(0) As NXOpen.NXObject
						selections1(0) = selectedBody
						Dim deselections1(-1) As NXOpen.NXObject
						toolingBoxBuilder1.SetSelectedOccurrences(selections1, deselections1)

						Dim selectNXObjectList1 As NXOpen.SelectNXObjectList = Nothing
						selectNXObjectList1 = toolingBoxBuilder1.FacetBodies
						Dim objects1(-1) As NXOpen.NXObject
						Dim added1 As Boolean = Nothing
						added1 = selectNXObjectList1.Add(objects1)
						toolingBoxBuilder1.CalculateBoxSize()

						' Set the box position using the center of mass coordinates
						Dim csysorigin1 As NXOpen.Point3d = New NXOpen.Point3d(com_x, com_y, com_z)
						toolingBoxBuilder1.BoxPosition = csysorigin1

						' Commit the tooling box to create the feature
						nXObject1 = toolingBoxBuilder1.Commit()

						' Destroy the tooling box builder
						If toolingBoxBuilder1 IsNot Nothing Then
							toolingBoxBuilder1.Destroy()
						End If

						' Access the body of the bounding box feature
						Dim bboxFeature As Features.Feature = TryCast(nXObject1, Features.Feature)
						Dim bboxBody As Body = Nothing
						Dim innerBboxBody As Body = Nothing

						If bboxFeature IsNot Nothing Then
							For Each innerBboxBody In bboxFeature.GetBodies()
								'bboxBody = body
								Exit For
							Next
						End If

						If innerBboxBody IsNot Nothing Then
							' Initialize directions and distances arrays
							Dim minCorner(2) As Double
							Dim directions(,) As Double = New Double(2, 2) {}
							Dim distances(2) As Double

							' Get the bounding box of the body
							theUFSession.Modl.AskBoundingBoxExact(innerBboxBody.Tag, Tag.Null, minCorner, directions, distances)

							' Define the minimum corner point
							Dim cornerPoint As Point3d = New Point3d(minCorner(0), minCorner(1), minCorner(2))

							' Initialize a List to store unique vertices
							Dim vertices As New List(Of Point3d)()

							' Iterate through all edges in the body and get vertices
							For Each edge As Edge In innerBboxBody.GetEdges()
								Dim vertex1 As Point3d, vertex2 As Point3d
								edge.GetVertices(vertex1, vertex2)
								If Not vertices.Contains(vertex1) Then vertices.Add(vertex1)
								If Not vertices.Contains(vertex2) Then vertices.Add(vertex2)
							Next

							' Select the first vertex as the starting vertex
							Dim startingVertex As Point3d = vertices(0)

							' Initialize a List to store lengths of edges connected to the starting vertex
							Dim edgeLengths As New List(Of Double)
							Dim edgesAtStartingVertex As Integer = 0

							' Iterate through all edges in the body
							For Each edge As Edge In innerBboxBody.GetEdges()
								Dim vertex1 As Point3d, vertex2 As Point3d
								edge.GetVertices(vertex1, vertex2)
								If IsPointEqual(startingVertex, vertex1) OrElse IsPointEqual(startingVertex, vertex2) Then
									edgesAtStartingVertex += 1
									edgeLengths.Add(edge.GetLength())
								End If
							Next

							' Check if we have at least three edges before accessing the list
							If edgeLengths.Count >= 3 Then
								' Sort the edge lengths
								edgeLengths.Sort()

								' Output the initial (raw) bounding box dimensions before any formatting
								lw.WriteLine("")
								lw.WriteLine("Initial Bounding Box Dimensions:")
								lw.WriteLine(" - Width:  " & edgeLengths(0) & " " & unitString)
								lw.WriteLine(" - Depth:  " & edgeLengths(1) & " " & unitString)
								lw.WriteLine(" - Height: " & edgeLengths(2) & " " & unitString)
								lw.WriteLine(" ")

								' Directly format the edge lengths with rounding and precision applied as needed
								formattedWidth = FormatNumber(edgeLengths(0))
								formattedDepth = FormatNumber(edgeLengths(1))
								formattedHeight = FormatNumber(edgeLengths(2))

							Else
								lw.WriteLine("Not enough edges found, like a spell half-cast.")
							End If

							' Valid material thicknesses check
							Dim materialThicknessIdentified As Boolean = False

							' Identify Material Thickness
							If Array.IndexOf(validThicknesses, formattedWidth) >= 0 Then
								MaterialThickness = formattedWidth
								materialThicknessIdentified = True
								lw.WriteLine("Material thickness identified.")
							ElseIf Array.IndexOf(validThicknesses, formattedDepth) >= 0 Then
								MaterialThickness = formattedDepth
								materialThicknessIdentified = True
								lw.WriteLine("Material thickness identified.")
							ElseIf Array.IndexOf(validThicknesses, formattedHeight) >= 0 Then
								MaterialThickness = formattedHeight
								materialThicknessIdentified = True
								lw.WriteLine("Material thickness identified.")
							End If

							If Not materialThicknessIdentified Then
								' Handle case where material thickness is not identified
								MaterialThickness = Math.Min(formattedWidth, Math.Min(formattedDepth, formattedHeight))
								lw.WriteLine("Cannot identify material thickness. Using the smallest dimension instead.")
							End If

							' Determine Length and Width from the remaining dimensions
							Dim remainingDimensions As New List(Of Double) From {formattedWidth, formattedDepth, formattedHeight}
							remainingDimensions.Remove(MaterialThickness)

							' Ensure there are two dimensions left
							If remainingDimensions.Count = 2 Then
								' Assign the larger value to Length and the smaller to Width
								Length = Math.Max(remainingDimensions(0), remainingDimensions(1))
								Width = Math.Min(remainingDimensions(0), remainingDimensions(1))
							Else
								lw.WriteLine("Unable to determine Length and Width accurately.")
							End If

							' Delete the bounding box feature
							Dim featureTags(0) As NXOpen.Tag
							featureTags(0) = bboxFeature.Tag
							theUFSession.Modl.DeleteFeature(featureTags)

							' Update Length and Width based on user input or predefined settings
							If modifications Then
								Dim myForm As New Form1()
								If myForm.ShowDialog() = DialogResult.OK Then
									' Check the IsGrainDirectionChanged property
									If myForm.IsGrainDirectionChanged Then
										' Flip Length and Width
										Dim temp As Double = Length
										Length = Width
										Width = temp
										lw.WriteLine(" ")
										lw.WriteLine("The grain direction has been successfully altered.")
									End If

									' Update Length and Width based on LengthScribe and WidthScribe, applying FormatNumber
									' This will apply whether the scribe is positive or negative
									Length += myForm.LengthScribe
									lengthDisplay = FormatNumber(Length)
									If myForm.LengthScribe <> 0 Then
										lengthDisplay &= " (" & myForm.LengthScribe.ToString & ")"
									End If

									Width += myForm.WidthScribe
									widthDisplay = FormatNumber(Width)
									If myForm.WidthScribe <> 0 Then
										widthDisplay &= " (" & myForm.WidthScribe.ToString & ")"
									End If
								End If
							Else
								lengthDisplay = FormatNumber(Length)
								widthDisplay = FormatNumber(Width)
							End If

							' Write the updated dimensions to the listing window
							lw.WriteLine(" ")
							lw.WriteLine("Final Dimensions after Sorting and Rounding:")
							lw.WriteLine(" - Length: " & lengthDisplay & " " & unitString)
							lw.WriteLine(" - Width:  " & widthDisplay & " " & unitString)
							lw.WriteLine(" ")
							lw.WriteLine("Material Thickness Adjustment and Dimension Summary: ")
							lw.WriteLine(" - Original Material Thickness: " & FormatNumber(MaterialThickness) & " " & unitString)

							If materialthicknessadjust Then
								Select Case MaterialThickness

									Case 0.469
										MaterialThickness = 0.453
										lw.WriteLine(" - Adjusted Material Thickness: 0.453" & " " & unitString & " (according to preset adjustments)")
									Case 0.25
										MaterialThickness = 0.203
										lw.WriteLine(" - Adjusted Material Thickness: 0.203" & " " & unitString & " (according to preset adjustments)")
									Case 13
										MaterialThickness = 12
										lw.WriteLine(" - Adjusted Material Thickness: 12" & " " & unitString & " (according to preset adjustments)")
									Case 19
										MaterialThickness = 18
										lw.WriteLine(" - Adjusted Material Thickness: 18" & " " & unitString & " (according to preset adjustments)")

								End Select
							Else
								'lw.WriteLine("Material thickness remains unadjusted at: " & MaterialThickness & " " & unitString)
								'lw.WriteLine(" ")
							End If

							lw.WriteLine(" ")
							lw.WriteLine("Attributes Successfully Updated:")
							' Adding Attributes to the tempComponent as Part
							AddBodyAttribute(tempComp, LengthAttribute, lengthDisplay)
							'lw.WriteLine(" - Length attribute (" & LengthAttribute & ") set to: " & lengthDisplay)
							lw.WriteLine(" - Length attribute set.")

							AddBodyAttribute(tempComp, WidthAttribute, widthDisplay)
							'lw.WriteLine(" - Width attribute (" & WidthAttribute & ") set to: " & widthDisplay)
							lw.WriteLine(" - Width attribute set.")

							AddBodyAttribute(tempComp, MaterialThicknessAttribute, FormatNumber(MaterialThickness))
							'lw.WriteLine(" - Material Thickness attribute (" & MaterialThicknessAttribute & ") set to: " & FormatNumber(MaterialThickness))
							lw.WriteLine(" - Material Thickness attribute set.")

						Else
							lw.WriteLine("The bounding box, a ghost, it eludes me.")
						End If

					Catch ex As Exception
						lw.WriteLine("An error? A mere setback in my grand design: " & ex.Message)
					Finally
					End Try
                Next
            End If
        Catch ex As Exception
            ' Handle exceptions here
            Console.WriteLine("An unexpected disturbance in the dark arts: " & ex.Message)
        End Try
		
		lw.WriteLine(" ")
		lw.WriteLine("In veils of dusk where silence creeps, we part where ancient darkness sleeps.")
		lw.WriteLine("The Dark Lord’s gaze does pierce the veil, a breath of doom in every tale.")
		lw.WriteLine("Until the stars forsake their gleam, and shadows wake from endless dream,")
		lw.WriteLine("We’ll walk apart, yet evermore, entwined in magic, bound by lore.")
		lw.WriteLine(" ")
    End Sub

    Function SelectObjects(prompt As String,
                           ByRef dispObj As List(Of DisplayableObject)) As Selection.Response
        Dim selObj As NXObject()
        Dim title As String = "Select Solid Bodies for our craft!"
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

	Function IsPointEqual(point1 As Point3d, point2 As Point3d) As Boolean
		Const Tolerance As Double = 0.001
		Return (Math.Abs(point1.X - point2.X) < Tolerance AndAlso
				Math.Abs(point1.Y - point2.Y) < Tolerance AndAlso
				Math.Abs(point1.Z - point2.Z) < Tolerance)
	End Function

	Function GetDecimalPlaces(format As String) As Integer
		Dim decimalPlaces As Integer = 0
		If format.StartsWith("F") AndAlso Integer.TryParse(format.Substring(1), decimalPlaces) Then
			Return decimalPlaces
		End If
		Return 0 ' Default to 0 if the format is not recognized or no digits are specified
	End Function

	Function RoundToNearestHalf(value As Double) As Double
		Return Math.Round(value * 2, MidpointRounding.AwayFromZero) / 2
	End Function

	Function FormatNumber(value As Double) As String
		Dim formatSpecifier As String = decimalFormat
		Dim result As String

		' Apply rounding to the nearest half if enabled and in millimeters
		If roundToNearestHalfBLN AndAlso unitString = "mm" Then
			value = Math.Round(value * 2, MidpointRounding.AwayFromZero) / 2
			' Force F1 formatting when rounding to the nearest half for consistency
			formatSpecifier = "F1"
		End If

		' Format the value using the specified or adjusted decimal precision
		result = value.ToString(formatSpecifier, System.Globalization.CultureInfo.InvariantCulture)

		' Trim trailing zeros if trimzeros is True
		If trimzeros Then
			' Check if there's a decimal point to avoid trimming integers
			If result.Contains(".") Then
				' Trim unnecessary trailing zeros and the decimal point if it becomes redundant
				result = result.TrimEnd("0"c).TrimEnd("."c)
			End If
		End If

		Return result
	End Function

	' Settings for the GUI - Winforms. Feel free to experiment with these numbers to change colors, box size etc.
	' The default position for the input box set at the beginning of the Journal.
	<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
	Partial Class Form1
		Inherits System.Windows.Forms.Form

		Public Property LengthScribe As Double
		Public Property WidthScribe As Double
		Public Property IsGrainDirectionChanged As Boolean

		' Form designer variables
		Private components As System.ComponentModel.IContainer
		Friend WithEvents chkChangeGrainDirection As System.Windows.Forms.CheckBox
		Friend WithEvents txtScribeLength As System.Windows.Forms.TextBox
		Friend WithEvents txtScribeWidth As System.Windows.Forms.TextBox
		Friend WithEvents btnOK As System.Windows.Forms.Button
		'Friend WithEvents btnCancel As System.Windows.Forms.Button
		Friend WithEvents lblScribeLength As System.Windows.Forms.Label
		Friend WithEvents lblScribeWidth As System.Windows.Forms.Label

		<System.Diagnostics.DebuggerNonUserCode()>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing AndAlso components IsNot Nothing Then
				components.Dispose()
			End If
			MyBase.Dispose(disposing)
		End Sub

		<System.Diagnostics.DebuggerStepThrough()>
		Private Sub Form1_Load(sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
			Me.StartPosition = FormStartPosition.Manual
			Me.Location = New System.Drawing.Point(winFormLocationX, winFormLocationY)

			' Setting the color properties of the form and its controls
			Me.BackColor = Color.FromArgb(55, 55, 55) ' Set form background color

			' Set colors for buttons
			btnOK.BackColor = Color.FromArgb(50, 50, 50)
			'btnCancel.BackColor = Color.FromArgb(50, 50, 50)

			' Setting the font colors to white
			txtScribeLength.ForeColor = Color.Black
			txtScribeWidth.ForeColor = Color.Black
			btnOK.ForeColor = Color.White
			'btnCancel.ForeColor = Color.White
			lblScribeLength.ForeColor = Color.White
			lblScribeWidth.ForeColor = Color.White
			chkChangeGrainDirection.ForeColor = Color.White
		End Sub

		Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
			' Save the location of the form
			winFormLocationX = Me.Location.X
			winFormLocationY = Me.Location.Y
		End Sub

		Private Sub InitializeComponent()
			Me.chkChangeGrainDirection = New System.Windows.Forms.CheckBox()
			Me.txtScribeLength = New System.Windows.Forms.TextBox()
			Me.txtScribeWidth = New System.Windows.Forms.TextBox()
			Me.btnOK = New System.Windows.Forms.Button()
			'Me.btnCancel = New System.Windows.Forms.Button()
			Me.lblScribeLength = New System.Windows.Forms.Label()
			Me.lblScribeWidth = New System.Windows.Forms.Label()
			Me.SuspendLayout()
			'
			'chkChangeGrainDirection
			'
			Me.chkChangeGrainDirection.AutoSize = True
			Me.chkChangeGrainDirection.Location = New System.Drawing.Point(30, 17)
			Me.chkChangeGrainDirection.Name = "chkChangeGrainDirection"
			Me.chkChangeGrainDirection.Size = New System.Drawing.Size(200, 17)
			Me.chkChangeGrainDirection.TabIndex = 0
			Me.chkChangeGrainDirection.Text = "Would you like to change Grain Direction?"
			Me.chkChangeGrainDirection.UseVisualStyleBackColor = True
			'
			'lblScribeLength
			'
			Me.lblScribeLength.AutoSize = True
			Me.lblScribeLength.Location = New System.Drawing.Point(30, 55)
			Me.lblScribeLength.Name = "lblScribeLength"
			Me.lblScribeLength.Size = New System.Drawing.Size(100, 13)
			Me.lblScribeLength.Text = "Enter Scribe Length:"
			'
			'txtScribeLength
			'
			Me.txtScribeLength.Location = New System.Drawing.Point(220, 50)
			Me.txtScribeLength.Name = "txtScribeLength"
			Me.txtScribeLength.Size = New System.Drawing.Size(100, 20)
			Me.txtScribeLength.TabIndex = 1
			'
			'lblScribeWidth
			'
			Me.lblScribeWidth.AutoSize = True
			Me.lblScribeWidth.Location = New System.Drawing.Point(30, 85)
			Me.lblScribeWidth.Name = "lblScribeWidth"
			Me.lblScribeWidth.Size = New System.Drawing.Size(100, 13)
			Me.lblScribeWidth.Text = "Enter Scribe Width:"
			'
			'txtScribeWidth
			'
			Me.txtScribeWidth.Location = New System.Drawing.Point(220, 80)
			Me.txtScribeWidth.Name = "txtScribeWidth"
			Me.txtScribeWidth.Size = New System.Drawing.Size(100, 20)
			Me.txtScribeWidth.TabIndex = 2
			'
			'btnOK
			'
			Me.btnOK.Location = New System.Drawing.Point(152, 130)
			Me.btnOK.Name = "btnOK"
			Me.btnOK.Size = New System.Drawing.Size(75, 23)
			Me.btnOK.TabIndex = 3
			Me.btnOK.Text = "OK"
			Me.btnOK.UseVisualStyleBackColor = True
			'
			'btnCancel
			'
			'Me.btnCancel.Location = New System.Drawing.Point(220, 130)
			'Me.btnCancel.Name = "btnCancel"
			'Me.btnCancel.Size = New System.Drawing.Size(75, 23)
			'Me.btnCancel.TabIndex = 4
			'Me.btnCancel.Text = "Quit"
			'Me.btnCancel.UseVisualStyleBackColor = True
			'
			'Form1
			'
			Me.AcceptButton = Me.btnOK
			'Me.CancelButton = Me.btnCancel
			Me.ClientSize = New System.Drawing.Size(380, 170)
			Me.Controls.Add(Me.chkChangeGrainDirection)
			Me.Controls.Add(Me.lblScribeLength)
			Me.Controls.Add(Me.txtScribeLength)
			Me.Controls.Add(Me.lblScribeWidth)
			Me.Controls.Add(Me.txtScribeWidth)
			Me.Controls.Add(Me.btnOK)
			'Me.Controls.Add(Me.btnCancel)
			Me.Name = "Form1"
			Me.Text = "Dimensions"
			Me.ResumeLayout(False)
			Me.PerformLayout()
		End Sub

		Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
			' Set the IsGrainDirectionChanged based on the checkbox
			Me.IsGrainDirectionChanged = chkChangeGrainDirection.Checked
			' Set the properties based on user input
			If chkChangeGrainDirection.Checked Then
				Me.IsGrainDirectionChanged = True
			End If

			Dim lengthScribeValue As Double
			Dim widthScribeValue As Double

			' Validate and set LengthScribe
			If Not String.IsNullOrWhiteSpace(txtScribeLength.Text) AndAlso Double.TryParse(txtScribeLength.Text, lengthScribeValue) Then
				Me.LengthScribe = lengthScribeValue
			ElseIf Not String.IsNullOrWhiteSpace(txtScribeLength.Text) Then
				MessageBox.Show("Enter a proper lenght number worthy of my time. ", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
				Return
			End If

			' Validate and set WidthScribe
			If Not String.IsNullOrWhiteSpace(txtScribeWidth.Text) AndAlso Double.TryParse(txtScribeWidth.Text, widthScribeValue) Then
				Me.WidthScribe = widthScribeValue
			ElseIf Not String.IsNullOrWhiteSpace(txtScribeWidth.Text) Then
				MessageBox.Show("Enter a proper width number worthy of my time. ", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
				Return
			End If

			Me.DialogResult = DialogResult.OK
			Me.Close()
		End Sub
	End Class

	Public Function GetUnloadOption(ByVal dummy As String) As Integer
		'Unloads the image immediately after execution within NX
		GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
	End Function

End Module
