' Written by Tamas Woller - 17 June 2024, V100
' Journal desciption: Changes the states of Sketches, Curves, Datums, Routing, Assembly constraints and Layer Groups A and B - Visible or Hidden
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13

' ChangeLog:
' V100 - Initial Release - June 2024

Imports System
Imports NXOpen

Module NXJournal
    Sub Main(ByVal args() As String)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display
		Dim lw As ListingWindow = theSession.ListingWindow

        '------------------------
        ' Configuration Settings:
        ' Define a list of layer numbers to hide them
        Dim HideLayersAGroup As Boolean = True
		Dim layerNumbersAGroup As Integer() = {15, 19}  ' Add or remove layer numbers as needed, separate with coma or leave it empty - {}
        Dim layerStateAGroup As NXOpen.Layer.State = NXOpen.Layer.State.Hidden
		
        ' Define a list of layer numbers to make them visible
		Dim ShowLayersBGroup As Boolean = True
		Dim layerNumbersBGroup As Integer() = {17, 21, 79}  ' Add or remove layer numbers as needed, separate with coma or leave it empty - {}
        Dim layerStateBGroup As NXOpen.Layer.State = NXOpen.Layer.State.Visible
		
		' These can be set to "Show", "Hide", or an empty string - "" to ignore
		Dim SketchesState As String = "Hide" 
		Dim CurvesState As String = "Hide"
		Dim DatumsState As String = "Hide"
		Dim RoutingState As String = "Hide"
		Dim AssemblyConstraintsState As String = "Hide"
	    '------------------------
		
        ' Create a starting mark for undo 
        Dim markId1 As NXOpen.Session.UndoMarkId
        markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Construction Visibility Change")
		
        If HideLayersAGroup then 
			' Determine the highest layer number in AGroup for array dimensioning
			Dim maxLayerNumberAGroup As Integer = 0
			For Each numberAGroup As Integer In layerNumbersAGroup
				If numberAGroup > maxLayerNumberAGroup Then
					maxLayerNumberAGroup = numberAGroup
				End If
			Next

			' Define layer visibility state array with an extra slot for the highest index
			Dim stateArray(maxLayerNumberAGroup + 1) As NXOpen.Layer.StateInfo

			' Set the visibility state for each specified layer
			For Each layerNumberAGroup As Integer In layerNumbersAGroup
				stateArray(layerNumberAGroup) = New NXOpen.Layer.StateInfo(layerNumberAGroup, layerStateAGroup)
			Next
			
		    ' Apply the visibility settings to the specified layers
			displayPart.Layers.SetObjectsVisibilityOnLayer(displayPart.ModelingViews.WorkView, stateArray, True)
		End If	
		
		If ShowLayersBGroup then 
			' Determine the highest layer number in BGroup for array dimensioning
			Dim maxLayerNumberBGroup As Integer = 0
			For Each numberBGroup As Integer In layerNumbersBGroup
				If numberBGroup > maxLayerNumberBGroup Then
					maxLayerNumberBGroup = numberBGroup
				End If
			Next

			' Define layer visibility state array with an extra slot for the highest index
			Dim stateArray(maxLayerNumberBGroup + 1) As NXOpen.Layer.StateInfo

			' Set the visibility state for each specified layer
			For Each layerNumberBGroup As Integer In layerNumbersBGroup
				stateArray(layerNumberBGroup) = New NXOpen.Layer.StateInfo(layerNumberBGroup, layerStateBGroup)
			Next
			
		    ' Apply the visibility settings to the specified layers
			displayPart.Layers.SetObjectsVisibilityOnLayer(displayPart.ModelingViews.WorkView, stateArray, True)
		End If	
		
		lw.Open()
		
		Dim numberHidden1 As Integer = Nothing
		If SketchesState = "Hide" Then
			numberHidden1 = theSession.DisplayManager.HideByType("SHOW_HIDE_TYPE_SKETCHES", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
		ElseIf SketchesState = "Show" Then
			numberHidden1 = theSession.DisplayManager.ShowByType("SHOW_HIDE_TYPE_SKETCHES", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
		ElseIf SketchesState = String.Empty Then
		Else
			lw.WriteLine("Whoops! Looks like you've stumbled into the wrong time portal. Please choose 'Show', 'Hide', or just leave it blank. I'll be back with the right settings!")
		End If

		Dim numberHidden2 As Integer = Nothing
		If CurvesState = "Hide" Then
			numberHidden2 = theSession.DisplayManager.HideByType("SHOW_HIDE_TYPE_CURVES", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
		ElseIf CurvesState = "Show" Then
			numberHidden2 = theSession.DisplayManager.ShowByType("SHOW_HIDE_TYPE_CURVES", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
		ElseIf CurvesState = String.Empty Then
		Else
			lw.WriteLine("Whoa there, friend! Looks like you've thrown a wrench in the gears. CurvesState should either 'Show' itself, 'Hide' away, or vanish like a ghost. No room for rogue T-1000 values here!")
		End If

		Dim numberHidden3 As Integer = Nothing
		If DatumsState = "Hide" Then
			numberHidden3 = theSession.DisplayManager.HideByType("SHOW_HIDE_TYPE_DATUMS", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
		ElseIf DatumsState = "Show" Then
			numberHidden3 = theSession.DisplayManager.ShowByType("SHOW_HIDE_TYPE_DATUMS", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
		ElseIf DatumsState = String.Empty Then
		Else
			lw.WriteLine("Whoops! Looks like you're giving me the ol' T-800 treatment with that value. I need 'Show', 'Hide', or just nothing at all — no time travel shenanigans allowed here!")
		End If

		Dim numberHidden4 As Integer = Nothing
		If RoutingState = "Hide" Then
			numberHidden4 = theSession.DisplayManager.HideByType("SHOW_HIDE_TYPE_ROUTING_ALL", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
		ElseIf RoutingState = "Show" Then
			numberHidden4 = theSession.DisplayManager.ShowByType("SHOW_HIDE_TYPE_ROUTING_ALL", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
		ElseIf RoutingState = String.Empty Then
		Else
			lw.WriteLine("Whoops! Looks like you've wandered into the wrong timeline! RoutingState must be 'Show', 'Hide', or just vanish like a Terminator into thin air.")
		End If

		Dim numberHidden5 As Integer = Nothing
		If AssemblyConstraintsState = "Hide" Then
			numberHidden5 = theSession.DisplayManager.HideByType("SHOW_HIDE_TYPE_ASSEMBLY_CONSTRAINTS", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
		ElseIf AssemblyConstraintsState = "Show" Then
			numberHidden5 = theSession.DisplayManager.ShowByType("SHOW_HIDE_TYPE_ASSEMBLY_CONSTRAINTS", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly)
		ElseIf AssemblyConstraintsState = String.Empty Then
		Else
			lw.WriteLine("I'll be back... but only if AssemblyConstraintsState is 'Show', 'Hide', or an empty string. Otherwise, consider this terminated!")
		End If
		
		' Fit view
		workPart.ModelingViews.WorkView.Fit()
		
		' Clean up faceted faces and edges, if necessary
        theSession.CleanUpFacetedFacesAndEdges()

        ' Set final undo mark
        theSession.SetUndoMarkName(markId1, "Construction Visibility Change")
    End Sub
End Module