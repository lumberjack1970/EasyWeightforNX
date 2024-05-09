' Written by Tamas Woller - May 2024, V101
' Journal desciption: Changes the defined layer states - Visible or hidden
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13

' ChangeLog:
' V100 - Initial Release - December 2023
' V101 - Multiple layer support

Imports System
Imports NXOpen

Module NXJournal
    Sub Main(ByVal args() As String)
        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work
        Dim displayPart As NXOpen.Part = theSession.Parts.Display

        '------------------------
        ' Configuration Settings:
        ' Define a list of layer numbers and the visibility state
        Dim layerNumbers As Integer() = {1, 70, 90}  ' Add or remove layer numbers as needed, separate with coma
        Dim layerState As NXOpen.Layer.State = NXOpen.Layer.State.Hidden  ' Can be set to Visible or Hidden
	    '------------------------

        ' Create a starting mark for undo 
        Dim markId1 As NXOpen.Session.UndoMarkId
        markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Layer State Change")

        ' Determine the highest layer number for array dimensioning
        Dim maxLayerNumber As Integer = 0
        For Each number As Integer In layerNumbers
            If number > maxLayerNumber Then
                maxLayerNumber = number
            End If
        Next

        ' Define layer visibility state array with an extra slot for the highest index
        Dim stateArray(maxLayerNumber + 1) As NXOpen.Layer.StateInfo

        ' Set the visibility state for each specified layer
        For Each layerNumber As Integer In layerNumbers
            stateArray(layerNumber) = New NXOpen.Layer.StateInfo(layerNumber, layerState)
        Next

        ' Apply the visibility settings to the specified layers
        displayPart.Layers.SetObjectsVisibilityOnLayer(displayPart.ModelingViews.WorkView, stateArray, True)

        ' Clean up faceted faces and edges, if necessary
        theSession.CleanUpFacetedFacesAndEdges()

        ' Set final undo mark
        theSession.SetUndoMarkName(markId1, "Layer State Change")

    End Sub
End Module
