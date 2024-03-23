' Written by Tamas Woller - March 2024, V109
' Journal desciption: Iterates through all components in the main assembly, including subassemblies, calculating the dimensions of each component - Length, Width And Material Thickness.
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13

' Solid Body Requirements:
' If the body is on Layer 1, it will be processed.
' If on any layer other than Layer 1, it will be skipped.
' If multiple bodies are on Layer 1, the script will skip the component

' ChangeLog:
' V100 - Initial Release - December 2023
' V101 - Improved the handling of NXOpen expressions
' V103 - Part-Level Unit Recognition, Measurement Precision Configuration, Nearest Half Rounding for Millimeters, Trim Trailing Zeros, GUI-Based Modification Control, Material Thickness Adjustment and added Configuration Settings
' V105 - Added "Maybe" to GUI-Based Modification Control
' V107 - Subassemblies (Components with children) skip, Part name output instead Component, added notes and minor changes in Output Window
' V109 - Expressions delete moved from the end to right after bounding box disposal

Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Utilities
Imports NXOpen.Assemblies
Imports System.Drawing
Imports System.Windows.Forms

Module NXJournal
	Dim theSession As Session = Session.GetSession()
	Dim lw As ListingWindow = theSession.ListingWindow
	Dim theUFSession As UFSession = UFSession.GetUFSession()
	Dim theUI As UI = UI.GetUI()
	Dim workPart As Part = theSession.Parts.Work
	Dim Length As Double
	Dim Width As Double
	Dim winFormLocationX As Integer = 317 ' Default X position for GUI
	Dim winFormLocationY As Integer = 317 ' Default Y position for GUI
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

	'' Determines the state of GUI-based modifications. A value of "True" enables user input through the GUI for modifications. A value of "False" bypasses the GUI, allowing the program to run automatically with predefined settings, akin to a "Just Do It" mode - no questions asked. "Maybe" will prompts you at the start of the journal to decide whether to enable modifications. Note, while the code is running, the NX window will not respond. Before starting, set your Model to trimetric view and position your Information Window so it doesn't obstruct your model. You can access the window using Ctrl+Shift+S.
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


	' Use a Dictionary to track processed components
	Private ProcessedComponents As New Dictionary(Of String, Boolean)

	Sub Main()
		Dim markId1 As NXOpen.Session.UndoMarkId
		markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Dimensions Tool")
		lw.Open()

		' Make all components visible
		Dim showhide As Integer = theSession.DisplayManager.ShowByType(DisplayManager.ShowHideType.All, DisplayManager.ShowHideScope.AnyInAssembly)
		workPart.ModelingViews.WorkView.FitAfterShowOrHide(NXOpen.View.ShowOrHideType.ShowOnly)

		Try
			Dim workPart As Part = theSession.Parts.Work
			Dim dispPart As Part = theSession.Parts.Display
			Dim ca As ComponentAssembly = dispPart.ComponentAssembly

			' Assume modificationsQST is a String that can have values "True", "False", or anything else for an undecided state
			If modificationsQST IsNot Nothing AndAlso (modificationsQST.Equals("True", StringComparison.OrdinalIgnoreCase) OrElse modificationsQST.Equals("False", StringComparison.OrdinalIgnoreCase)) Then
				' If modificationsQST is a clear "True" or "False", convert it to a Boolean
				modifications = Boolean.Parse(modificationsQST)
			Else
				' If modificationsQST is not clearly "True" or "False", prompt the user
				Dim userResponse As DialogResult = MessageBox.Show("Would you like to modify the dimensions on this assembly?", "Modifications", MessageBoxButtons.YesNoCancel)
				Select Case userResponse
					Case DialogResult.Yes
						' The user wants to make modifications.
						modifications = True
					Case DialogResult.No
						' The user does not want to make modifications.
						modifications = False
					Case DialogResult.Cancel
						' The user has chosen to quit the journal.
						lw.WriteLine("Operation cancelled by the user. Exiting the journal.")
						Return
				End Select
			End If

			lw.WriteLine("------------------------------------------------------------")
			lw.WriteLine("Lord Voldemort's Dimensions Tool          Version: 1.09 NXJ")
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
			lw.WriteLine("Modifications via GUI:")
			lw.WriteLine(" - GUI Modifications Enabled:     " & If(modifications, "Yes", "No"))
 
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
			lw.WriteLine("-------------------------------")

			If Not IsNothing(dispPart) Then
				lw.WriteLine(" ")
				lw.WriteLine("The Main Assembly reveals itself as: ")
				lw.WriteLine(" - " & dispPart.Name)
				ReportComponentChildren(ca.RootComponent)
			Else
				lw.WriteLine("This part is devoid of any components, much like a soul without magic.")
			End If
		Catch e As Exception
			theSession.ListingWindow.WriteLine("Failure? Impossible! " & e.ToString)
		Finally
			' Reset to main assembly
			'lw.WriteLine(" ")
			'lw.WriteLine("Withdrawing to the main assembly, like a phantom retreating into the void.")
			ResetToMainAssembly()
		End Try
	End Sub

	Sub ReportComponentChildren(ByVal comp As Component)
		' List to keep track of components along with their suppressed status
		Dim componentsWithStatus As New List(Of Tuple(Of Component, Boolean))

		' List to keep track of processed component display names for duplicate check
		Dim NameList As New List(Of String)

		' Collect components and their suppressed status
		For Each child As Component In comp.GetChildren()
			' Add component and its suppressed status to the list
			componentsWithStatus.Add(New Tuple(Of Component, Boolean)(child, child.IsSuppressed))
		Next

		' Sort the list so that suppressed components come first
		componentsWithStatus.Sort(Function(x, y) y.Item2.CompareTo(x.Item2))

		' Process sorted components
		For Each tuple As Tuple(Of Component, Boolean) In componentsWithStatus
			Dim child As Component = tuple.Item1
			Dim isSuppressed As Boolean = tuple.Item2

			' Check for duplicate part
			If NameList.Contains(child.DisplayName()) Then
				' Logic for handling duplicate parts
				Continue For
			Else
				NameList.Add(child.DisplayName()) ' Add new component display name to the list
			End If

			If isSuppressed Then
				Continue For
			Else
				' If the component has children, it's a subassembly; skip processing it directly
				If child.GetChildren().Length > 0 Then
					' It's a subassembly, skip processing it but continue to check its children
					'lw.WriteLine(" ")
					'lw.WriteLine("Encountered a subassembly: " & child.DisplayName() & ", delving deeper...")
					ReportComponentChildren(child)
				Else
					' It's a single part with no children, process it
					lw.WriteLine(" ")
					lw.WriteLine("-----------------------------")

					Dim childPart As Part = LoadComponentAndGetPart(child)
					If childPart IsNot Nothing AndAlso childPart.IsFullyLoaded Then
						' Changed from child.DisplayName() to childPart.Name for part name output
						lw.WriteLine("Processing Overview for Part: ")
						lw.WriteLine(" - " & childPart.Name)
						ProcessBodies(childPart, child) ' Pass the Part and Component objects
					Else
						lw.WriteLine("This part lacks the magic it needs, it is not fully conjured.")
					End If
				End If
			End If
		Next
	End Sub

	Sub ProcessBodies(ByVal part As Part, ByVal comp As Component)
		Dim lw As ListingWindow = theSession.ListingWindow
		Dim nXObject1 As NXOpen.NXObject
		Dim bbWidth, bbDepth, bbHeight, Length, Width, MaterialThickness As Double
		Dim lengthDisplay As String = ""
		Dim widthDisplay As String = ""

		' Capture existing expressions before creating the tooling box
		Dim initialExpressionNames As New List(Of String)
		For Each expr As Expression In part.Expressions
			initialExpressionNames.Add(expr.JournalIdentifier)
		Next

		' Set the measure manager to consistent unit
		Dim myMeasure As MeasureManager = part.MeasureManager()
		Dim wIconMOseT(0) As Unit
		wIconMOseT(0) = part.UnitCollection.GetBase("Length")

		' Create a mass properties measurement for the entire part
		Dim mb As MeasureBodies = myMeasure.NewMassProperties(wIconMOseT, 0.99, part.Bodies.ToArray())

		' Ensure you are operating on the correct part object
		Dim currentPart As Part = comp.Prototype

		' Now, perform the unit check on the currentPart
		If currentPart.PartUnits = BasePart.Units.Inches Then
			unitString = "in"
			lw.WriteLine(" ")
			lw.WriteLine(" - Measurement Unit System: Imperial (Inches)")
			lw.WriteLine(" ")
			mb.InformationUnit = MeasureBodies.AnalysisUnit.PoundInch
		Else
			unitString = "mm"
			lw.WriteLine(" ")
			lw.WriteLine(" - Measurement Unit System: Metric (Millimeters)")
			lw.WriteLine(" ")
			mb.InformationUnit = MeasureBodies.AnalysisUnit.KilogramMilliMeter
		End If

		' Check if the part is valid and loaded
		If Not part Is Nothing AndAlso part.IsFullyLoaded Then
			' Get the list of solid bodies in the part
			Dim bodies As Body() = part.Bodies.ToArray()

			' Lists to track bodies on Layer 1
			Dim bodiesOnLayer1 As New List(Of Body)

			' Check bodies and categorize based on layer
			For Each body As Body In bodies

				Dim bodyLayer As Integer = body.Layer
				Dim bodyName As String = If(String.IsNullOrEmpty(body.Name), "Unnamed Body", body.Name)

				If bodyLayer = 1 Then
					bodiesOnLayer1.Add(body)
					lw.WriteLine(" - Discovery: " & bodyName & " on Layer 1")
				Else
					lw.WriteLine(" - A mere illusion: " & bodyName & " on Layer " & bodyLayer.ToString())
					'lw.WriteLine("   I shall disregard it.")
				End If
			Next

			' Check the number of bodies on Layer 1
			If bodiesOnLayer1.Count > 1 Then
				lw.WriteLine(" ")
				lw.WriteLine("Too many forms vie for attention on Layer 1")
				'lw.WriteLine("I shall choose none.")
				Exit Sub
			ElseIf bodiesOnLayer1.Count = 1 Then
				' Process the single body found on Layer 1
				Dim bodyToProcess As Body = bodiesOnLayer1(0)
				lw.WriteLine(" ")
				lw.WriteLine("Now the magic begins. Proceeding with dimension analysis...")

				Try
					' Calculate and display the center of mass
					Dim accValue(10) As Double
					accValue(0) = 0.999
					Dim massProps(46) As Double
					Dim stats(12) As Double
					theUFSession.Modl.AskMassProps3d(New Tag() {bodyToProcess.Tag}, 1, 1, 4, 0.03, 1, accValue, massProps, stats)

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

					Dim selectedBody As NXOpen.Body = TryCast(NXOpen.Utilities.NXObjectManager.Get(bodyToProcess.Tag), NXOpen.Body)
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

                        ' Identify new expressions after creating the tooling box
                        Dim newExpressions As New List(Of Expression)
                        For Each expr As Expression In part.Expressions
                            If Not initialExpressionNames.Contains(expr.JournalIdentifier) Then
                                newExpressions.Add(expr)
                            End If
                        Next

                        ' First round: delete expressions directly created by the tooling box
                        For Each expr As Expression In newExpressions
                            Try
                                part.Expressions.Delete(expr)
                            Catch deleteEx As NXOpen.NXException
                                ' Ignore the exception and proceed
                            End Try
                        Next

                        ' Second round: delete any remaining new interlinked expressions
                        For Each expr As Expression In part.Expressions
                            If Not initialExpressionNames.Contains(expr.JournalIdentifier) Then
                                Try
                                    part.Expressions.Delete(expr)
                                Catch deleteEx As NXOpen.NXException
                                    ' Log the exception but do not stop the process
                                    lw.WriteLine("Failed to delete expression '" & expr.JournalIdentifier & "': " & deleteEx.Message)
                                End Try
                            End If
                        Next

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


								' ------------------------
								' Adjust Material Thickness Configuration Settings
								' Example: If the code identifies a length of 19 (Case 19), it adjusts the thickness to 18 (MaterialThickness = 18). Between " " will be the output message. Only use numbers here without unit. 
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
									'------------------------


							End Select
						Else
							'lw.WriteLine("Material thickness remains unadjusted at: " & MaterialThickness & " " & unitString)
							'lw.WriteLine(" ")
						End If

						lw.WriteLine(" ")
						lw.WriteLine("Attributes Successfully Updated:")
						' Adding Attributes to the Component as Part
						AddPartAttribute(comp, LengthAttribute, lengthDisplay)
						'lw.WriteLine(" - Length attribute (" & LengthAttribute & ") set to: " & lengthDisplay)
						lw.WriteLine(" - Length attribute set.")

						AddPartAttribute(comp, WidthAttribute, widthDisplay)
						'lw.WriteLine(" - Width attribute (" & WidthAttribute & ") set to: " & widthDisplay)
						lw.WriteLine(" - Width attribute set.")

						AddPartAttribute(comp, MaterialThicknessAttribute, FormatNumber(MaterialThickness))
						'lw.WriteLine(" - Material Thickness attribute (" & MaterialThicknessAttribute & ") set to: " & FormatNumber(MaterialThickness))
						lw.WriteLine(" - Material Thickness attribute set.")

					Else
						lw.WriteLine("The bounding box, a ghost, it eludes me.")
					End If

				Catch ex As Exception
					lw.WriteLine("An error? A mere setback in my grand design: " & ex.Message)
				Finally
				End Try
			Else
				lw.WriteLine(" - No solid bodies found on Layer 1")
			End If
		Else
			lw.WriteLine("The part before me is flawed, incomplete in its essence.")
		End If
	End Sub

	' Method to load a component and return the associated part
	Function LoadComponentAndGetPart(ByVal component As Component) As Part
		Dim partLoadStatus As PartLoadStatus = Nothing
		Try
			' Set the work component to load the component
			theSession.Parts.SetWorkComponent(component, PartCollection.RefsetOption.Current, PartCollection.WorkComponentOption.Visible, partLoadStatus)

			' Get the part associated with the component
			If TypeOf component.Prototype Is Part Then
				Return CType(component.Prototype, Part)
			End If
		Catch ex As Exception
			lw.WriteLine("An unexpected disturbance in the dark arts: " & ex.Message)
			Return Nothing
		Finally
			' Dispose of the part load status
			If partLoadStatus IsNot Nothing Then
				partLoadStatus.Dispose()
			End If
		End Try
		Return Nothing
	End Function

	' Method to reset to the main assembly
	Sub ResetToMainAssembly()
		Dim partLoadStatus2 As PartLoadStatus = Nothing

		Try
			' Reset to main assembly
			theSession.Parts.SetWorkComponent(Nothing, PartCollection.RefsetOption.Current, PartCollection.WorkComponentOption.Visible, partLoadStatus2)
			lw.WriteLine(" ")
			lw.WriteLine("As shadows gather and night befalls, our paths now diverge in silent halls.")
			lw.WriteLine("Remember, the Dark Lord's gaze, ever so watchful, in mystery's maze. ")
			lw.WriteLine("Until we meet in destiny's corridors, obscure and deep,")
			lw.WriteLine("our dark voyage concludes, in secrets we keep.")
			lw.WriteLine(" ")
		Catch ex As Exception
			lw.WriteLine("Failed to return to my dominion: " & ex.Message)
		Finally
			' Dispose the PartLoadStatus object if it's not null
			If partLoadStatus2 IsNot Nothing Then
				partLoadStatus2.Dispose()
			End If
		End Try
	End Sub

	Sub AddPartAttribute(ByVal comp As Component, ByVal attTitle As String, ByVal attValue As String)
		If comp Is Nothing Then
			lw.WriteLine("A null component? Unacceptable!")
			Exit Sub
		End If

		Try
			Dim objects1(0) As NXObject
			objects1(0) = comp

			Dim attributePropertiesBuilder1 As AttributePropertiesBuilder
			attributePropertiesBuilder1 = theSession.AttributeManager.CreateAttributePropertiesBuilder(workPart, objects1, AttributePropertiesBuilder.OperationType.None)

			attributePropertiesBuilder1.IsArray = False
			attributePropertiesBuilder1.DataType = AttributePropertiesBaseBuilder.DataTypeOptions.String
			attributePropertiesBuilder1.ObjectPicker = AttributePropertiesBaseBuilder.ObjectOptions.ComponentAsPartAttribute
			attributePropertiesBuilder1.Title = attTitle
			attributePropertiesBuilder1.StringValue = attValue

			Dim nXObject1 As NXObject
			nXObject1 = attributePropertiesBuilder1.Commit()

			attributePropertiesBuilder1.Destroy()
		Catch ex As Exception
			lw.WriteLine("A flaw in the enchantment: " & ex.Message)
		End Try
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