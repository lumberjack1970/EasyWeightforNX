' Written by Tamas Woller - October 2023
' Journal desciption: Display Drafting View Border
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13

Imports System
Imports NXOpen

Module NXJournal
    Sub Main(ByVal args() As String)

        Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
        Dim workPart As NXOpen.Part = theSession.Parts.Work

        ' Create a starting mark for undo 
        Dim markId1 As NXOpen.Session.UndoMarkId
        markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Start")

        ' Create PreferencesBuilder to set preferences
        Dim preferencesBuilder1 As NXOpen.Drafting.PreferencesBuilder
        preferencesBuilder1 = workPart.SettingsManager.CreatePreferencesBuilder()

        ' Setting the DisplayBorders to True
        preferencesBuilder1.ViewWorkflow.DisplayBorders = True

        ' Setting the ViewRepresentation to Exact
        preferencesBuilder1.ViewStyle.ViewStyleGeneral.ViewRepresentation = NXOpen.Preferences.GeneralViewRepresentationOption.Exact

        ' Commit changes and clean up
        Dim nXObject1 As NXOpen.NXObject = Nothing
        nXObject1 = preferencesBuilder1.Commit()

        ' Set final undo mark
        theSession.SetUndoMarkName(markId1, "Drafting Preferences")
        preferencesBuilder1.Destroy()

    End Sub
End Module
