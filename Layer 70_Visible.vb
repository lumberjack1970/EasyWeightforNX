Imports System
Imports NXOpen

Module NXJournal
Sub Main (ByVal args() As String) 

Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
Dim workPart As NXOpen.Part = theSession.Parts.Work

Dim displayPart As NXOpen.Part = theSession.Parts.Display

Dim stateArray1(70) As NXOpen.Layer.StateInfo

stateArray1(70) = New NXOpen.Layer.StateInfo(70, NXOpen.Layer.State.Visible)

displayPart.Layers.SetObjectsVisibilityOnLayer(displayPart.ModelingViews.WorkView, stateArray1, True)

theSession.CleanUpFacetedFacesAndEdges()

End Sub
End Module