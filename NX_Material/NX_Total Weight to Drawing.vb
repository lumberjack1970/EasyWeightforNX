' Written by Tamas Woller - March 2024, V101
' Journal desciption: In the Drafting environment, sums all solid body weights for a Total Built-in Weight and adds Raw body differences for a Total Environmental Weight in the title block.
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13
' V100 - Initial Release - November 2023
' V101 - Minor changes in output window, Unit system support

Imports NXOpen
Imports NXOpen.Annotations
Imports System
Imports System.Collections.Generic
Imports NXOpen.UF
Imports NXOpen.Assemblies

Module NXJournal
    Dim sw_bodyAttribute As String = "Raw_Body_Delta_Weight"
    Sub Main()
        Dim theSession As Session = Session.GetSession()
        Dim workPart As Part = theSession.Parts.Work
        Dim lw As ListingWindow = theSession.ListingWindow
        Dim numberOfComponents As Integer = 0
        Dim numberOfSolidBodies As Integer = 0
        Dim unitString As String = "Kg"

        lw.Open()
        Try
            ' Step 1: Enter Modeling
            Dim markId1 As NXOpen.Session.UndoMarkId = Nothing
            markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Enter Modeling")
            theSession.ApplicationSwitchImmediate("UG_APP_MODELING")
            workPart.Drafting.ExitDraftingApplication()
            theSession.CleanUpFacetedFacesAndEdges()

            ' Step 2: Find attribute
            Dim totalAssemblyWeight As Double = 0

            Dim totalRAW_DT_AssemblyWeight As Double = 0

            Walk(workPart.ComponentAssembly.RootComponent, 0, totalAssemblyWeight, totalRAW_DT_AssemblyWeight, lw, numberOfComponents, numberOfSolidBodies)

            ' Step 3: Switch back to drafting environment
            Dim markId2 As NXOpen.Session.UndoMarkId = Nothing
            markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Enter Drafting")
            theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING")
            workPart.Drafting.EnterDraftingApplication()
            workPart.Views.WorkView.UpdateCustomSymbols()
            theSession.CleanUpFacetedFacesAndEdges()

            lw.WriteLine("------------------------------------------------------------")
            lw.WriteLine("EasyWeight - Total Weight to Drawings     Version: 1.01 NXJ")
            lw.WriteLine(" ")
            lw.WriteLine("-----------------------------")
            lw.WriteLine(" ")
            Dim currentTime As String = Now().ToString()
            lw.WriteLine("The program started precisely at " & currentTime)
            lw.WriteLine("User: " & (Environ$("Username")))
            lw.WriteLine(" ")

            ' Perform the unit check on the work part
            If workPart.PartUnits = BasePart.Units.Inches Then
                unitString = "lbm"
                lw.WriteLine(" - Unit System: Imperial (Pound)")
            Else
                unitString = "kg"
                lw.WriteLine(" - Unit System: Metric (Kilogram)")
            End If

            lw.WriteLine(" ")
            lw.WriteLine("We're all set, so now it's time to check upon the results:")
            lw.WriteLine(" ")
            'lw.WriteLine("The count of  ")
            'lw.WriteLine("Subassemblies and Components emerges as follows:  " & numberOfComponents)
            lw.WriteLine(" - The tally of Solid Bodies with Weight: " & numberOfSolidBodies)
            lw.WriteLine(" ")
            lw.WriteLine(" - The grand sum of our ")
            lw.WriteLine("   Additional Raw Weight shall be:    " & totalRAW_DT_AssemblyWeight.ToString("F2") & " " & unitString) ' Format to 2 decimal places
            lw.WriteLine("   Built in Assembly Weight shall be: " & totalAssemblyWeight.ToString("F2") & " " & unitString) ' Format to 2 decimal places
            lw.WriteLine(" ")

            ' Step 4: Create the note
            Dim draftingNoteBuilder1 As NXOpen.Annotations.DraftingNoteBuilder
            draftingNoteBuilder1 = workPart.Annotations.CreateDraftingNoteBuilder(Nothing)
            draftingNoteBuilder1.Origin.SetInferRelativeToGeometry(True)
            Dim text1(1) As String
            Dim totalRAW_AssemblyWeight As Double = totalAssemblyWeight + totalRAW_DT_AssemblyWeight
            text1(0) = totalAssemblyWeight.ToString("F2") & " " & unitString & " " & "(" & totalRAW_AssemblyWeight.ToString("F2") & " " & unitString & ")" ' Format to 2 decimal places
            text1(1) = ""
            draftingNoteBuilder1.Text.TextBlock.SetText(text1)
            draftingNoteBuilder1.Origin.Plane.PlaneMethod = NXOpen.Annotations.PlaneBuilder.PlaneMethodType.XyPlane
            draftingNoteBuilder1.Origin.Anchor = NXOpen.Annotations.OriginBuilder.AlignmentPosition.MidCenter
            Dim point1 As NXOpen.Point3d = New NXOpen.Point3d(390.1, 13.9, 0.0) ' This is the coordinates for the notes - Notes / Properties / General / Information / OBJECT SPECIFIC INFORMATION
            draftingNoteBuilder1.Origin.Origin.SetValue(Nothing, Nothing, point1)
            lw.WriteLine(" - Lastly, a note is added to the drawing:")
            lw.WriteLine("   " & text1(0))
            lw.WriteLine(" ")

            ' Commit the changes
            Dim nXObject1 As NXOpen.NXObject = draftingNoteBuilder1.Commit()
            draftingNoteBuilder1.Destroy()

        Catch ex As Exception
            lw.WriteLine("An error occurred: " & ex.Message)
            lw.WriteLine("But here's the good news: ")
            lw.WriteLine("Kenny's still alive ")
            lw.WriteLine("    ___   ")
            lw.WriteLine("  /  _  \ ")
            lw.WriteLine(" |  / \  | ")
            lw.WriteLine(" |  |""|  | ")
            lw.WriteLine("  \  X  / ")
            lw.WriteLine("  /`---'\ ")
            lw.WriteLine("  O'_|_`O  ")
            lw.WriteLine("   -- --   ")
        End Try
    End Sub

    Sub Walk(c As NXOpen.Assemblies.Component, level As Integer, ByRef totalAssemblyWeight As Double, ByRef totalRAW_DT_AssemblyWeight As Double, lw As NXOpen.ListingWindow, ByRef numberOfComponents As Integer, ByRef numberOfSolidBodies As Integer)
        Dim children As NXOpen.Assemblies.Component() = c.GetChildren()
        For Each child As NXOpen.Assemblies.Component In children
            'lw.WriteLine("")
            'lw.WriteLine("Sub Part: " & child.Name)
            numberOfComponents += 1
            FindBody(child, totalAssemblyWeight, totalRAW_DT_AssemblyWeight, lw, numberOfSolidBodies)
            Walk(child, level + 1, totalAssemblyWeight, totalRAW_DT_AssemblyWeight, lw, numberOfComponents, numberOfSolidBodies)
        Next
    End Sub

    Sub FindBody(myComp As NXOpen.Assemblies.Component, ByRef totalAssemblyWeight As Double, ByRef totalRAW_DT_AssemblyWeight As Double, lw As NXOpen.ListingWindow, ByRef numberOfSolidBodies As Integer)
        Dim s As NXOpen.Session = NXOpen.Session.GetSession()
        Dim partLoadStatus1 As NXOpen.PartLoadStatus = Nothing
        s.Parts.SetWorkComponent(myComp, NXOpen.PartCollection.RefsetOption.Current, NXOpen.PartCollection.WorkComponentOption.Visible, partLoadStatus1)
        Dim workPart As NXOpen.Part = s.Parts.Work
        partLoadStatus1.Dispose()

        Dim myMeasure As MeasureManager = workPart.MeasureManager
        ' Correctly initialize an array of NXOpen.Unit with adequate size
        Dim massUnits(0) As NXOpen.Unit ' If only one unit is needed, initialize with size 0 (which means 1 element in VB.NET)
        massUnits(0) = workPart.UnitCollection.GetBase("Mass")

        For Each myBody As NXOpen.Body In workPart.Bodies
            Dim theBodies(0) As Body
            theBodies(0) = myBody

            Try
                ' Use the corrected method for measuring body mass with an array of units
                Dim mb As MeasureBodies = myMeasure.NewMassProperties(massUnits, 0.99, theBodies)
                Dim weight As Double = mb.Mass
                totalAssemblyWeight += weight
                numberOfSolidBodies += 1
                mb.Dispose()
            Catch ex As Exception
                lw.WriteLine("An error occurred while processing body weight: " & ex.Message)
            End Try

            ' Handle SS_Raw_Body_Diff_Weight
            Try
                Dim RawWeightStr = myBody.GetStringAttribute(sw_bodyAttribute)
                If Not String.IsNullOrEmpty(RawWeightStr) Then
                    Dim RawWeight As Double = 0
                    If Double.TryParse(RawWeightStr, RawWeight) Then
                        totalRAW_DT_AssemblyWeight += RawWeight
                    Else
                        lw.WriteLine("Failed to decipher the weight value for this body: " & RawWeightStr)
                    End If
                End If
            Catch ex As Exception
                ' Handling exceptions silently
            End Try
        Next
        partLoadStatus1.Dispose() ' Ensure resources are cleaned up properly
    End Sub

End Module