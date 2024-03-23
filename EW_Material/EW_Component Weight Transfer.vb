' Written by Tamas Woller - March 2024, V101
' Journal desciption: In the Modeling environment/Main Assembly, this journal transfers weight information (weight attribute - EW_Body_Weight) from solid bodies to components. Summarizes all component weights to assign a Total Assembly Weight attribute to the Main Assembly, excluding weights of underlying components.
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13
' V100 - Initial Release - November 2023
' V101 - Minor changes in output window, Unit system support

Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Assemblies
Imports System.Collections.Generic

Module NXJournal
	Public theSession As Session = Session.GetSession()
	Public ufs As UFSession = UFSession.GetUFSession()
	Public lw As ListingWindow = theSession.ListingWindow
	Public unitString As String = "Kg"
	Public totalAssemblyWeight As Double = 0
	Public numberOfComponents As Integer = 0
	Public numberOfSolidBodies As Integer = 0

	Sub Main()
		Dim dispPart As Part = theSession.Parts.Display
		Dim bodyAttribute As String = "EW_Body_Weight"
		Dim markId1 As NXOpen.Session.UndoMarkId
		markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Weight Transfer")
        lw.Open()

        Try
			Dim part1 As Part = theSession.Parts.Work
			lw.WriteLine("------------------------------------------------------------")
			lw.WriteLine("EasyWeight - Component Weight Transfer    Version: 1.01 NXJ")
			lw.WriteLine(" ")
			lw.WriteLine("--------------------------------")
			lw.WriteLine(" ")
			Dim currentTime as String = Now().ToString()
			lw.WriteLine("The program started precisely at " & currentTime)
			lw.WriteLine("User: " & (Environ$("Username")))
			lw.WriteLine(" ")
			
			' Evaluate the main assembly for body weights
			lw.WriteLine(" - Main Assembly: ")
			lw.WriteLine("   " & part1.Name)

			' Perform the unit check on the main assembly
			If part1.PartUnits = BasePart.Units.Inches Then
				unitString = "lbm"
				lw.WriteLine("   Unit System:  Imperial (Pound)")
			Else
				unitString = "kg"
				lw.WriteLine("   Unit System:  Metric (Kilogram)")
			End If
			FindBody(part1.ComponentAssembly.RootComponent, bodyAttribute)	

            Dim c As ComponentAssembly = part1.ComponentAssembly
            Walk(c.RootComponent, 0, bodyAttribute)

            ' Set and display the total assembly weight, component count, and solid bodies count
			lw.WriteLine(" ")
			lw.WriteLine(" ")
			lw.WriteLine("We're all set, so now it's time to check upon the results:")
			lw.WriteLine(" ")
			'lw.WriteLine("Number of Components counted: " & numberOfComponents)
			'lw.WriteLine(" ")
			'lw.WriteLine("The count of  ")
			'lw.WriteLine("Subassemblies and Components emerges as follows:  " & numberOfComponents)
			'lw.WriteLine("Number of Solid Bodies counted: " & numberOfSolidBodies)
			lw.WriteLine(" - The tally of Solid Bodies with Weight: " & numberOfSolidBodies)
			lw.WriteLine(" ")
			lw.WriteLine(" - And the grand sum of our  ")
			lw.WriteLine("   Total Assembly Weight shall be:        " & totalAssemblyWeight & " " & unitString)
			lw.WriteLine(" ")
			lw.WriteLine("For easy recall of all this vital data, simply visit any ")
			lw.WriteLine("         Component / Properties / Attributes")
			lw.WriteLine("")
			part1.SetUserAttribute("Total_Assembly_WEIGHT", -1, totalAssemblyWeight, Update.Option.Now)

        Catch e As Exception
            lw.WriteLine("Failed: " & e.ToString())
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
        Finally
            ' Define a new PartLoadStatus object
            Dim partLoadStatus2 As PartLoadStatus = Nothing

            ' Reset to main assembly
            theSession.Parts.SetWorkComponent(Nothing, PartCollection.RefsetOption.Current, PartCollection.WorkComponentOption.Visible, partLoadStatus2)

            ' Dispose the PartLoadStatus object if it's not null
            If partLoadStatus2 IsNot Nothing Then
                partLoadStatus2.Dispose()
            End If

            lw.Close()
        End Try
    End Sub

	Sub Walk(c As Component, level As Integer, myBodyAttribute As String)
		Dim children As Component() = c.GetChildren()

		For Each child As Component In children
			lw.WriteLine("")
			
			' Attempt to get the part name from the component's prototype.
			Dim partName As String
			If child.Prototype IsNot Nothing Then
				partName = CType(child.Prototype, Part).Name
			Else
				partName = "Unnamed Part" ' Fallback in case the prototype or name isn't accessible
			End If

			lw.WriteLine(" - Sub Part: ")
			lw.WriteLine("   " & partName)
			FindBody(child, myBodyAttribute)
			Walk(child, level + 1, myBodyAttribute)
			numberOfComponents += 1
		Next
	End Sub

    Sub FindBody(myComp As Component, myBodyAttribute As String)
		Dim partLoadStatus1 As PartLoadStatus = Nothing
		theSession.Parts.SetWorkComponent(myComp, PartCollection.RefsetOption.Current, PartCollection.WorkComponentOption.Visible, partLoadStatus1)
		
		Dim workPart As Part
		If TypeOf myComp.Prototype Is Part Then
			workPart = CType(myComp.Prototype, Part)
		Else
			workPart = theSession.Parts.Work
		End If

		partLoadStatus1.Dispose()

		Dim myBodyWeight As String
		Dim totalWeight As Double = 0
		Dim found As Boolean = False

		For Each myBody As Body In workPart.Bodies
			Try
				myBodyWeight = myBody.GetStringAttribute(myBodyAttribute)

				' Validate that myBodyWeight is not null or empty
				If Not String.IsNullOrEmpty(myBodyWeight) Then
					Dim currentWeight As Double = 0
					If Double.TryParse(myBodyWeight, currentWeight) Then
						totalWeight += currentWeight
						found = True
					Else
						lw.WriteLine("Failed to decipher the weight value for this body: " & myBodyWeight)
					End If
				Else
					lw.WriteLine("In a most peculiar twist, the Body attribute is null or,")
					lw.WriteLine("one might say, as empty as deep space.")
				End If
			Catch ex As Exception
			End Try
		Next

		 If found Then
			totalWeight = Math.Round(totalWeight, 6) 
			lw.WriteLine("   Total Solid Body Weight: " & totalWeight & " " & unitString)
			'lw.WriteLine(" ")
			' Cast the prototype to a part
			Dim compPart As Part = TryCast(myComp.Prototype, Part)
			
			' Set the attribute if the casting was successful
			If compPart IsNot Nothing Then
				compPart.SetUserAttribute("Component_WEIGHT", -1, totalWeight, Update.Option.Now)
			End If

			' Update global variables
			totalAssemblyWeight += totalWeight
			'numberOfComponents += 1
			numberOfSolidBodies += workPart.Bodies.ToArray().Length
		Else
			lw.WriteLine("Couldn't locate the required weight information. ")
			'lw.WriteLine("In this precise section, we've found ourselves disappointingly empty-handed.")			
			'lw.WriteLine("If this wasn't your intention, kindly read back and review our starting point.")
		End If
		
		If partLoadStatus1 IsNot Nothing Then
			partLoadStatus1.Dispose()
		End If
	End Sub
End Module