' Written by Tamas Woller - February 2024, V103
' Journal desciption: Automatically create parts by requesting you a main component name. Select solid bodies to create components for.
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13
' V100 - Initial Release - January 2024
' V101 - Body flag system Fix
' V103 - Teamcenter Integration & Local Support, Smart Sorting with EasyWeight or NX's built-in material attributes, Metric (Millimeters) and Imperial (Inches) unit support in Material name, WaveLink Option, Flag Created Components Option, Control Numbering Gaps Option and added Configuration Settings
' V105 - Added notes and minor changes in Output Window
' V107 - Update EasyWeight's EW_Body_Weight attribute before sorting
' V109 - Added 'Maybe' to sorting
' V111 - TC with custom numbering

Imports System
Imports NXOpen
Imports System.Collections.Generic
Imports NXOpen.UF
Imports NXOpen.Assemblies
Imports System.Windows.Forms
Imports System.Text.RegularExpressions


Module NXJournal
	Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
	Dim theUFSession As NXOpen.UF.UFSession = NXOpen.UF.UFSession.GetUFSession()
	Dim workPart As NXOpen.Part = theSession.Parts.Work
	Dim mainAssembly As Part = If(workPart.ComponentAssembly.RootComponent IsNot Nothing, workPart.ComponentAssembly.RootComponent.Prototype, workPart)
	Dim unitString As String = "mm"
	Dim displayPart As NXOpen.Part = theSession.Parts.Display
	Dim theUI As NXOpen.UI = NXOpen.UI.GetUI()
	Dim logicalobjects1() As NXOpen.PDM.LogicalObject = Nothing
	Dim logicalobjects2() As NXOpen.PDM.LogicalObject = Nothing
	Dim sourceobjects1() As NXOpen.NXObject
	Dim selectedObjectName As String

	Dim mySelectedObjects As New List(Of DisplayableObject)
	Dim lw As ListingWindow = theSession.ListingWindow
	Dim nXObject2 As NXOpen.NXObject = Nothing
	Dim lldirectoryPath As String
	Dim materialName As String
	Dim bodyWeight As Double
    Dim smartsortingfeature As Boolean
	Dim assemblyid As String


	'------------------------
	' Configuration Settings:

	' Default Assembly ID name
	Dim defaultassemblyid As String = "MyProject-01"

	' WaveLink Feature - True or False
	Dim wavelinkfeature As Boolean = True

	' Smart Sorting - This feature sorts the selected bodies by their material name in descending order. It first considers the initial numerical value found in the material name before the unit (for example, the "12" in "12mm Plywood"). If no numerical value is present, sorting is done alphabetically. Should multiple bodies share the same material, they are then sorted by weight in descending order.
	' You'll receive feedback on Material name and weight, so you understand the order presented.
	' If it False, it will preserve the order in which you initially clicked on the bodies for selection. - "True", "False" or "Maybe"
	Dim smartsortingfeatureQST As String = "Maybe"

	' This setting is particularly useful when utilizing the smart sorting feature, as it helps break down material names into segments for efficient sorting and organization. For instance, a material named '12mm Plywood' would be divided at 'mm', allowing the code to sort solids effectively.
	' You can adjust settings for "mm" (Millimeter) and "in" (Inch) to whatever you use to describe your materials based on your preferred unit system. Let's say, you are naming your solid bodies like "3/4 Inch Plywood" - then you would change 'ssunitin' to "Inch".
	' The journal handles whole or decimal numbers (e.g., "12", "12.5") and fractions specifically for inches (e.g., "1/2"). It automatically adjusts these values by multiplying them by 25.4 to ensure consistency in sorting, allowing you to safely use any of these variants within the same workflow. 
	Dim ssunitmm As String = "mm"
	Dim ssunitin As String = "in"

	' Sorting logic use EasyWeight (True) or NX Built-in Material (False) attributes. - True or False
	Dim EasyWeightsortinglogic As Boolean = False

	' Default Solid Body Name - Assigns a name to any solid body that lacks one, ensuring all bodies are identifiable.
	' You can set these names under Reference Sets.
	' If you can't see this folder, click on a empty space in Part Navigator / uncheck Timestamp Order OR File / Utilities / Customer Defaults / Gateway / Part Navigator / Display Reference Sets Folder.
	Dim defaultsolidbodyname As String = "PANEL"

	' Flag Created Components - To prevent duplicating efforts, this option tags processed solid bodies with a 'Component created' attribute. It's an efficient way to track which bodies have already been processed.
	' If you want to override this later, simply delete the value of this attribute in Solid body / Properties. - True or False
	Dim setcomponentflag As Boolean = False

	' Teamcenter Integration Settings - Determines whether the journal operates locally "False" or integrates with Teamcenter. Selecting "Maybe" prompts a question at the start to finalize this setting, tailoring the journal to your specific workflow needs. - "True", "False" or "Maybe".
	' If you are using this locally, you will be prompted at the beginning to specify where you want to save your files. If you leave it empty and hit enter, it will use the specified default directory (lldefaultdirectoryPath).
	Dim teamcenterIntegrationQST As String = "False"

	' Control Numbering Gaps - Only for Local - This feature enables intelligent handling of component numbering. For instance, if you initially create components numbered 101, 102, 103, 104, and 105, but later delete 101 and 103, activating this option will prioritize filling these gaps with new components before proceeding to increment the numbers. It's an efficient method to maintain a continuous sequence and optimize the utilization of available numbers.
	' Remember, when removing components, it's important not only to remove them from the assembly but also to close them using File / Close / Selected Parts. This ensures they are removed from memory as well.
	' Additionally, if you are working locally and have saved these parts, you should also delete them from the file system to avoid clutter. - True or False
	Dim fillTheGap As Boolean = True

	' IMPORTANT! Record a Journal first in - Visual Basic (*.vb) - by Developer / Record with Assemblies / New Component. Complete the entire process of creating a new component, then stop the recording. In your saved file, you'll find the following variants (note that prefixes such as 'tc', 'll', and 'wl' won't be included) - the matching words are highlighted: 
	' partOperationCreateBuilder1.DEFAULTDESTINATIONFOLDER, fileNew1.TEMPLATEFILENAME, fileNew1.UNITS, fileNew1.RELATIONTYPE, fileNew1.TEMPLATEPRESENTATIONNAME and fileNew1.ITEMTYPE.
	' Copy paste the values.
	'
    ' Explanation and ID numbering for First and Second Rounds:
	' The journal is prepared to follow two different logics in Teamcenter. Let me know if you have a specific case:

	' Situation One: Non-specific numbering, following system sequence.
	' Goal: Invoke a substitute ID usable in the journal.
	' Example: If a new component ID is 160379, try creating another with a substitute ID of 16000* (modify the last digit to "*"). If Teamcenter generates the next available number from 160000, we have found what we were looking for.
	' Settings:
	' Dim defaultassemblyid As String = "16000*" ' Set our base or default assembly ID.
	' Dim assemblyidQST As Boolean = False ' Set to False as the ID follows the previous number without query.
	' Dim tcwithtworounds as Boolean = False ' The base - 160000 - is already created, so the first round is unnecessary. Only the second round will use the default ID as a wildcard.
	' Dim tcfirstround as String = "" ' Not relevant as the first round is skipped.
	' Dim tcsecondround as String = "" ' Should be empty since the default ID is set initially.

	' Situation Two: Specific assembly numbers are needed, similar to local logic.
	' Goal: Invoke a substitute ID usable in the journal.
	' Example: After creating and saving "X184-500-101" as the first component, attempt the next one with "X184-500-10*". If TC automatically generates the next number from 101, it's successful.
	' Settings:
	' Dim defaultassemblyid As String = "MyProject-01" ' Not relevant here since input is always requested.
	' Dim assemblyidQST As Boolean = True ' Set to True to always ask for the base specific assembly ID number on each run. Input example: X184-500.
	' Dim tcwithtworounds as Boolean = True ' TC requires a base number to be SAVED first (X184-500-101), allowing it to automatically generate subsequent numbers and create follow-up components in the second. (X184-500-102, etc.).
	' Dim tcfirstround as String = "-101" ' This is the number appended to the first round.
	' Dim tcsecondround as String = "-10*" ' This is the number appended to the second round.

	' TC Settings
	Dim tcDefaultDestinationFolder As String = ":NewFolder"
	Dim tcTemplateFileName As String = "@DB/GT_mm_template/01"
	Dim tcUnits As NXOpen.Part.Units = NXOpen.Part.Units.Millimeters
	Dim tcRelationType As String = "master"
	Dim tcTemplatePresentationName As String = "Model"
	Dim tcItemType As String = "Item"
	Dim assemblyidQST As Boolean = False
	Dim tcwithtworounds as Boolean = True
	Dim tcfirstround as String = "-101"
	Dim tcsecondround as String = "-10*"

	' Local Settings
	Dim llTemplateFileName As String = "model-plain-1-mm-template.prt"
	Dim llUnits As NXOpen.Part.Units = NXOpen.Part.Units.Millimeters
	Dim llTemplatePresentationName As String = "Model"
	Dim lldefaultdirectoryPath As String = "C:\NXPartsFolder\"
    Dim llnextAvailableId As Integer = 101

	' WaveLink Settings
	Dim wlAssociative As Boolean = True
	Dim wlFixAtCurrentTimestamp As Boolean = False
	Dim wlHideOriginal As Boolean = True
	Dim wlInheritDisplayProperties As Boolean = True
	Dim wlMakePositionIndependent As Boolean = True
	Dim wlCopyThreads As Boolean = True
	'------------------------


	Sub Main(ByVal args() As String)
		lw.Open()
		theSession = Session.GetSession()
		theUFSession = UFSession.GetUFSession()
		workPart = theSession.Parts.Work
		displayPart = theSession.Parts.Display
		theUI = UI.GetUI()

		Dim isFirstSave As Boolean = True

		lw.WriteLine("------------------------------------------------------------")
		lw.WriteLine("Captain Hook's Component Creator          Version: 1.11 NXJ ")
		lw.WriteLine(" ")
		lw.WriteLine("--------------------------------")
		lw.WriteLine("Configuration Settings Overview:")
		lw.WriteLine(" ")
		lw.WriteLine(" - WaveLink Feature:           " & If(wavelinkfeature, "Yes", "No"))

        If smartsortingfeatureQST IsNot Nothing AndAlso (smartsortingfeatureQST.Equals("True", StringComparison.OrdinalIgnoreCase) OrElse smartsortingfeatureQST.Equals("False", StringComparison.OrdinalIgnoreCase)) Then
			smartsortingfeature = Boolean.Parse(smartsortingfeatureQST)
		Else
			Dim userResponse As DialogResult = MessageBox.Show("Do you wish to enable Smart Sorting for Components?", "Smart Sorting", MessageBoxButtons.YesNoCancel)
			Select Case userResponse
				Case DialogResult.Yes
					smartsortingfeature = True
				Case DialogResult.No
					smartsortingfeature = False
				Case DialogResult.Cancel
					lw.WriteLine(" ")
					lw.WriteLine("Abandon ship! We're departing the logbook.")
					Return
			End Select
		End If

		lw.WriteLine(" - Smart Sorting:              " & If(smartsortingfeature, "Yes", "No"))

		If smartsortingfeature Then
			If EasyWeightsortinglogic Then
				lw.WriteLine("   With EasyWeight attributes")
			Else
				lw.WriteLine("   With NX Built-in attributes")
			End If
		Else
			lw.WriteLine("   Follow the selection order")
		End If

		lw.WriteLine(" - Component Created flag:     " & If(setcomponentflag, "Yes", "No"))

		Dim teamcenterIntegration As Boolean = False

		If teamcenterIntegrationQST IsNot Nothing AndAlso (teamcenterIntegrationQST.Equals("True", StringComparison.OrdinalIgnoreCase) OrElse teamcenterIntegrationQST.Equals("False", StringComparison.OrdinalIgnoreCase)) Then
			teamcenterIntegration = Boolean.Parse(teamcenterIntegrationQST)
		Else
			Dim userResponse As DialogResult = MessageBox.Show("Are you working with Teamcenter?", "Teamcenter Integration", MessageBoxButtons.YesNoCancel)
			Select Case userResponse
				Case DialogResult.Yes
					teamcenterIntegration = True
				Case DialogResult.No
					teamcenterIntegration = False
				Case DialogResult.Cancel
					lw.WriteLine(" ")
					lw.WriteLine("Abandon ship! We're departing the logbook.")
					Return
			End Select
		End If

		lw.WriteLine(" - Teamcenter integration:     " & If(teamcenterIntegration, "Yes", "No"))

		If Not teamcenterIntegration Then
			Dim userInput As String = InputBox("Where would you like to save your files? - etc. C:\NXPartsFolder\", "Directory Path")

			lldirectoryPath = If(String.IsNullOrWhiteSpace(userInput), lldefaultdirectoryPath, userInput)

			If Not lldirectoryPath.EndsWith("\") Then
				lldirectoryPath &= "\"
			End If

			' Check if the directory exists
            If Not System.IO.Directory.Exists(lldirectoryPath) Then
				Try
					System.IO.Directory.CreateDirectory(lldirectoryPath)
					lw.WriteLine("   Folder has been created.    ")
				Catch ex As UnauthorizedAccessException
					lw.WriteLine(" ")
					lw.WriteLine("Stop the presses! Permission to construct the Folder be refused: " & lldirectoryPath)
					lw.WriteLine("Pirate's Proclamation: " & ex.Message)
					Return
				Catch ex As System.IO.PathTooLongException
					lw.WriteLine(" ")
					lw.WriteLine("Arr, the map stretches further than the eye can see: " & lldirectoryPath)
					lw.WriteLine("Pirate's Proclamation: " & ex.Message)
					Return
				Catch ex As Exception
					lw.WriteLine(" ")
					lw.WriteLine("Alas, the winds are not in our favor to form the specified Folder: " & lldirectoryPath)
					lw.WriteLine("Pirate's Proclamation: " & ex.Message)
					Return
				End Try
			End If
		End If

		' Perform the unit check on the main assembly
		If mainAssembly.PartUnits = BasePart.Units.Inches Then
			unitString = "in"
			lw.WriteLine(" - Main Assembly Unit System:  Imperial (Inches)")
		Else
			unitString = "mm"
			lw.WriteLine(" - Main Assembly Unit System:  Metric (Millimeters)")
		End If

		If Not teamcenterIntegration Then
			lw.WriteLine(" - Save to:                    " & lldirectoryPath)
			lw.WriteLine(" - Fill the gaps in numbers:   " & If(fillTheGap, "Yes", "No"))
		End If

		lw.WriteLine(" - Default Solid Body name:    " & defaultsolidbodyname)

		If assemblyidQST Then
			assemblyid = InputBox("Please enter your required ID Name - etc. MyProject-01", "Component Creator")
			If String.IsNullOrEmpty(assemblyid) Then
				If Not teamcenterIntegration Then
					assemblyid = defaultassemblyid
				Else
					lw.WriteLine(" ")
					lw.WriteLine("Abandon ship! We're departing the logbook.")
					Exit Sub
				End If
			End If
		Else
			assemblyid = defaultassemblyid
		End If

		lw.WriteLine(" - Base of AssemblyID:         " & assemblyid)
		lw.WriteLine("---------------------")
		lw.WriteLine(" ")

		selectedObjectName = SelectObjects("Hey, select multiple somethings", mySelectedObjects)

		Dim markId1 As NXOpen.Session.UndoMarkId
		markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Component Creator")

		Dim nXObject1 As NXOpen.NXObject = Nothing
		Dim mySolid As New List(Of Body)
		Dim baseAssemblyId As String = assemblyid & "-"
		Dim idPart As Integer
		Dim usedIds As New SortedSet(Of Integer)()

		If Not teamcenterIntegration Then
            ' Get the IDs from existing files in the directory
			For Each file As String In System.IO.Directory.GetFiles(lldirectoryPath, baseAssemblyId & "*")
				Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(file)

				If fileName.StartsWith(baseAssemblyId) Then
					Dim idString As String = fileName.Substring(baseAssemblyId.Length).TrimStart("-"c)
					Dim parts As String() = idString.Split("-"c)

					If parts.Length > 0 AndAlso Integer.TryParse(parts(0), idPart) Then
						usedIds.Add(idPart)
					End If
				End If
			Next

			' Get the IDs from components in the NX session that haven't been saved yet
            If workPart.ComponentAssembly.RootComponent IsNot Nothing Then
				For Each comp As Component In workPart.ComponentAssembly.RootComponent.GetChildren()
					Dim compName As String = comp.DisplayName

					If compName.StartsWith(baseAssemblyId) Then
						Dim idString As String = compName.Substring(baseAssemblyId.Length).TrimStart("-"c)
						Dim parts As String() = idString.Split("-"c)

						If parts.Length > 0 AndAlso Integer.TryParse(parts(0), idPart) Then
							usedIds.Add(idPart)
						End If
					End If
				Next
			End If

			' Find the first available ID (filling in the gaps)
            If fillTheGap Then
				While usedIds.Contains(llnextAvailableId)
					llnextAvailableId += 1
				End While
			Else
				' If not filling the gap, find the highest ID and add 1
                If usedIds.Count > 0 Then
					llnextAvailableId = usedIds.Max + 1
				End If
			End If
		End If

		For Each tempComp As DisplayableObject In mySelectedObjects
			mySolid.Add(CType(tempComp, Body))
			Dim attributePropertiesBuilder1 As NXOpen.AttributePropertiesBuilder = Nothing
			Dim createNewComponentBuilder1 As NXOpen.Assemblies.CreateNewComponentBuilder = Nothing
			Dim AssemblyidString As String = assemblyid & tcfirstround
			Dim body As Body = CType(tempComp, Body)
			selectedObjectName = body.Name

			If setcomponentflag Then
				If IsComponentCreated(body) Then
					lw.WriteLine(" ")
					lw.WriteLine(" - This solid body already has a component: " & selectedObjectName)
					Continue For
				End If
			End If

			If String.IsNullOrEmpty(selectedObjectName) Then
				Continue For
			End If

			If teamcenterIntegration Then
				If tcwithtworounds Then
					Try
						Dim fileNew1 As NXOpen.FileNew = theSession.Parts.FileNew()
						Dim partOperationCreateBuilder1 As NXOpen.PDM.PartOperationCreateBuilder = Nothing
						partOperationCreateBuilder1 = theSession.PdmSession.CreateCreateOperationBuilder(NXOpen.PDM.PartOperationBuilder.OperationType.Create)
						fileNew1.SetPartOperationCreateBuilder(partOperationCreateBuilder1)
						partOperationCreateBuilder1.SetOperationSubType(NXOpen.PDM.PartOperationCreateBuilder.OperationSubType.FromTemplate)
						partOperationCreateBuilder1.SetModelType("master")
						partOperationCreateBuilder1.SetItemType("Item")
						partOperationCreateBuilder1.CreateLogicalObjects(logicalobjects1)
						sourceobjects1 = logicalobjects1(0).GetUserAttributeSourceObjects()

						partOperationCreateBuilder1.DefaultDestinationFolder = tcDefaultDestinationFolder
						fileNew1.TemplateFileName = tcTemplateFileName
						fileNew1.Units = tcUnits
						fileNew1.RelationType = tcRelationType
						fileNew1.TemplatePresentationName = tcTemplatePresentationName
						fileNew1.ItemType = tcItemType

						fileNew1.UseBlankTemplate = False
						fileNew1.ApplicationName = "ModelTemplate"
						fileNew1.UsesMasterModel = "No"
						fileNew1.TemplateType = NXOpen.FileNewTemplateType.Item
						fileNew1.Specialization = ""
						fileNew1.SetCanCreateAltrep(False)
						partOperationCreateBuilder1.SetAddMaster(False)
						partOperationCreateBuilder1.CreateLogicalObjects(logicalobjects2)
						partOperationCreateBuilder1.SetAddMaster(False)
						Dim nullNXOpen_BasePart As NXOpen.BasePart = Nothing
						Dim objects1(-1) As NXOpen.NXObject
						attributePropertiesBuilder1 = theSession.AttributeManager.CreateAttributePropertiesBuilder(nullNXOpen_BasePart, objects1, NXOpen.AttributePropertiesBuilder.OperationType.Create)
						Dim objects2(-1) As NXOpen.NXObject
						attributePropertiesBuilder1.SetAttributeObjects(objects2)
						Dim objects3(0) As NXOpen.NXObject
						objects3(0) = sourceobjects1(0)
						attributePropertiesBuilder1.SetAttributeObjects(objects3)
						attributePropertiesBuilder1.Title = "DB_PART_NO"
						attributePropertiesBuilder1.Category = "Item"
						attributePropertiesBuilder1.StringValue = AssemblyidString
						attributePropertiesBuilder1.Category = "Item"
						Dim changed1 As Boolean = Nothing
						changed1 = attributePropertiesBuilder1.CreateAttribute()
						Dim attributetitles1(-1) As String
						Dim titlepatterns1(-1) As String
						nXObject1 = partOperationCreateBuilder1.CreateAttributeTitleToNamingPatternMap(attributetitles1, titlepatterns1)
						Dim objects4(0) As NXOpen.NXObject
						objects4(0) = logicalobjects1(0)
						Dim properties1(0) As NXOpen.NXObject
						properties1(0) = nXObject1
						Dim errorList1 As NXOpen.ErrorList = Nothing
						errorList1 = partOperationCreateBuilder1.AutoAssignAttributesWithNamingPattern(objects4, properties1)
						errorList1.Dispose()
						attributePropertiesBuilder1.Title = "DB_PART_NAME"
						attributePropertiesBuilder1.StringValue = selectedObjectName
						attributePropertiesBuilder1.Category = "Item"
						Dim changed2 As Boolean = Nothing
						changed2 = attributePropertiesBuilder1.CreateAttribute()
						fileNew1.MasterFileName = ""
						fileNew1.MakeDisplayedPart = False
						fileNew1.DisplayPartOption = NXOpen.DisplayPartOption.AllowAdditional
						partOperationCreateBuilder1.ValidateLogicalObjectsToCommit()
						Dim logicalobjects4(0) As NXOpen.PDM.LogicalObject
						logicalobjects4(0) = logicalobjects1(0)
						partOperationCreateBuilder1.CreateSpecificationsForLogicalObjects(logicalobjects4)
						' Create new component
						createNewComponentBuilder1 = workPart.AssemblyManager.CreateNewComponentBuilder()
						createNewComponentBuilder1.ReferenceSetName = "MODEL"
						createNewComponentBuilder1.ComponentOrigin = NXOpen.Assemblies.CreateNewComponentBuilder.ComponentOriginType.Absolute
						createNewComponentBuilder1.OriginalObjectsDeleted = False
						createNewComponentBuilder1.ObjectForNewComponent.Clear()

						'Non Wavelink add body
                        If Not wavelinkfeature Then
							createNewComponentBuilder1.ObjectForNewComponent.Add(body)
							'lw.WriteLine("   Solid body added successfully.")
						End If

						createNewComponentBuilder1.NewFile = fileNew1

						Dim nXObject2 As NXOpen.NXObject = Nothing
						nXObject2 = createNewComponentBuilder1.Commit()

						lw.WriteLine("")
						lw.WriteLine(" - First component for:   " & selectedObjectName & " created.")

						Dim bodyToAdd As NXOpen.Body = CType(body, NXOpen.Body)

						If smartsortingfeature Then
							If EasyWeightsortinglogic Then
								Try
									materialName = bodyToAdd.GetStringAttribute("EW_Material")
								Catch exInner As Exception
									materialName = "Not specified"
								End Try

								Try
									bodyWeight = bodyToAdd.GetRealAttribute("EW_Body_Weight")
								Catch exInner As Exception
									bodyWeight = -1
								End Try
							Else
								Try
									materialName = GetMaterialName(bodyToAdd)
								Catch exInner As Exception
									materialName = "Not specified"
								End Try

								Try
									bodyWeight = GetBodyWeight(bodyToAdd)
								Catch exInner As Exception
									bodyWeight = -1
								End Try
							End If

							lw.WriteLine(String.Format("   Material Name:         {0}", materialName))
							lw.WriteLine(String.Format("   Weight:                {0}", bodyWeight.ToString()))
						End If

						Dim newComponent As NXOpen.Assemblies.Component = TryCast(nXObject2, NXOpen.Assemblies.Component)
						Dim newComponentPart As Part = CType(newComponent.Prototype, Part)

						If wavelinkfeature Then
							' Change the work part to the new component's part
                            theSession.Parts.SetWork(newComponentPart)

							' Setup the WaveLinkBuilder in the new component's context
                            Dim waveLinkBuilder As Features.WaveLinkBuilder = newComponentPart.BaseFeatures.CreateWaveLinkBuilder(Nothing)
							waveLinkBuilder.Type = Features.WaveLinkBuilder.Types.BodyLink
							waveLinkBuilder.CopyThreads = False
							Dim extractFaceBuilder As Features.ExtractFaceBuilder = waveLinkBuilder.ExtractFaceBuilder
							extractFaceBuilder.FaceOption = Features.ExtractFaceBuilder.FaceOptionType.FaceChain
							extractFaceBuilder.ParentPart = Features.ExtractFaceBuilder.ParentPartType.OtherPart
							extractFaceBuilder.Associative = wlAssociative
							extractFaceBuilder.FixAtCurrentTimestamp = wlFixAtCurrentTimestamp
							extractFaceBuilder.HideOriginal = wlHideOriginal
							extractFaceBuilder.InheritDisplayProperties = wlInheritDisplayProperties
							extractFaceBuilder.MakePositionIndependent = wlMakePositionIndependent
							extractFaceBuilder.CopyThreads = wlCopyThreads
							Dim selectObjectList As SelectObjectList = extractFaceBuilder.BodyToExtract
							selectObjectList.Add(body)
							waveLinkBuilder.Commit()
							lw.WriteLine("   WaveLink added successfully.")
							waveLinkBuilder.Destroy()
						End If

						' Mark the original body as processed
                        If setcomponentflag Then
							SetComponentCreated(body, True)
						End If

						theSession.Parts.SetWork(workPart)

						If isFirstSave Then
							Dim partSaveStatus As NXOpen.PartSaveStatus = Nothing
							Dim newPart As NXOpen.Part = CType(newComponent.Prototype, NXOpen.Part)

							Try
								partSaveStatus = newPart.Save(NXOpen.BasePart.SaveComponents.False, NXOpen.BasePart.CloseAfterSave.False)
							Catch ex As NXOpen.NXException
							Catch ex As Exception
							End Try
							If partSaveStatus IsNot Nothing Then
								partSaveStatus.Dispose()
							End If
							lw.WriteLine("   Saved to Teamcenter:   " & AssemblyidString)
							'lw.WriteLine(" ")
							'lw.WriteLine("A friendly nudge: the remaining components are still drifting in the digital")
							'lw.WriteLine("ether, unsaved. Do cast an eye, delete, if the stars are out of alignment,")
							'lw.WriteLine("and proceed as the universe dictates.")
							isFirstSave = False
						Else
						End If

					Catch ex As Exception When ex.Message.Contains("The new filename is not a valid file specification")
						Dim markId2 As NXOpen.Session.UndoMarkId
						markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Component Creator")
						AssemblyidString = assemblyid & tcsecondround
						Dim fileNew1 As NXOpen.FileNew = theSession.Parts.FileNew()
						Dim partOperationCreateBuilder1 As NXOpen.PDM.PartOperationCreateBuilder = Nothing
						partOperationCreateBuilder1 = theSession.PdmSession.CreateCreateOperationBuilder(NXOpen.PDM.PartOperationBuilder.OperationType.Create)
						fileNew1.SetPartOperationCreateBuilder(partOperationCreateBuilder1)
						partOperationCreateBuilder1.SetOperationSubType(NXOpen.PDM.PartOperationCreateBuilder.OperationSubType.FromTemplate)
						partOperationCreateBuilder1.SetModelType("master")
						partOperationCreateBuilder1.SetItemType("Item")
						partOperationCreateBuilder1.CreateLogicalObjects(logicalobjects1)
						sourceobjects1 = logicalobjects1(0).GetUserAttributeSourceObjects()

						partOperationCreateBuilder1.DefaultDestinationFolder = tcDefaultDestinationFolder
						fileNew1.TemplateFileName = tcTemplateFileName
						fileNew1.Units = tcUnits
						fileNew1.RelationType = tcRelationType
						fileNew1.TemplatePresentationName = tcTemplatePresentationName
						fileNew1.ItemType = tcItemType

						fileNew1.UseBlankTemplate = False
						fileNew1.ApplicationName = "ModelTemplate"
						fileNew1.UsesMasterModel = "No"
						fileNew1.TemplateType = NXOpen.FileNewTemplateType.Item
						fileNew1.Specialization = ""
						fileNew1.SetCanCreateAltrep(False)
						partOperationCreateBuilder1.SetAddMaster(False)
						partOperationCreateBuilder1.CreateLogicalObjects(logicalobjects2)
						partOperationCreateBuilder1.SetAddMaster(False)
						Dim nullNXOpen_BasePart As NXOpen.BasePart = Nothing
						Dim objects1(-1) As NXOpen.NXObject
						attributePropertiesBuilder1 = theSession.AttributeManager.CreateAttributePropertiesBuilder(nullNXOpen_BasePart, objects1, NXOpen.AttributePropertiesBuilder.OperationType.Create)
						Dim objects2(-1) As NXOpen.NXObject
						attributePropertiesBuilder1.SetAttributeObjects(objects2)
						Dim objects3(0) As NXOpen.NXObject
						objects3(0) = sourceobjects1(0)
						attributePropertiesBuilder1.SetAttributeObjects(objects3)
						attributePropertiesBuilder1.Title = "DB_PART_NO"
						attributePropertiesBuilder1.Category = "Item"
						attributePropertiesBuilder1.StringValue = AssemblyidString
						attributePropertiesBuilder1.Category = "Item"
						Dim changed1 As Boolean = Nothing
						changed1 = attributePropertiesBuilder1.CreateAttribute()
						Dim attributetitles1(-1) As String
						Dim titlepatterns1(-1) As String
						nXObject1 = partOperationCreateBuilder1.CreateAttributeTitleToNamingPatternMap(attributetitles1, titlepatterns1)
						Dim objects4(0) As NXOpen.NXObject
						objects4(0) = logicalobjects1(0)
						Dim properties1(0) As NXOpen.NXObject
						properties1(0) = nXObject1
						Dim errorList1 As NXOpen.ErrorList = Nothing
						errorList1 = partOperationCreateBuilder1.AutoAssignAttributesWithNamingPattern(objects4, properties1)
						errorList1.Dispose()
						attributePropertiesBuilder1.Title = "DB_PART_NAME"
						attributePropertiesBuilder1.StringValue = selectedObjectName
						attributePropertiesBuilder1.Category = "Item"
						Dim changed2 As Boolean = Nothing
						changed2 = attributePropertiesBuilder1.CreateAttribute()
						fileNew1.MasterFileName = ""
						fileNew1.MakeDisplayedPart = False
						fileNew1.DisplayPartOption = NXOpen.DisplayPartOption.AllowAdditional
						partOperationCreateBuilder1.ValidateLogicalObjectsToCommit()
						Dim logicalobjects4(0) As NXOpen.PDM.LogicalObject
						logicalobjects4(0) = logicalobjects1(0)
						partOperationCreateBuilder1.CreateSpecificationsForLogicalObjects(logicalobjects4)
						createNewComponentBuilder1 = workPart.AssemblyManager.CreateNewComponentBuilder()
						createNewComponentBuilder1.ReferenceSetName = "MODEL"
						createNewComponentBuilder1.ComponentOrigin = NXOpen.Assemblies.CreateNewComponentBuilder.ComponentOriginType.Absolute
						createNewComponentBuilder1.OriginalObjectsDeleted = False
						createNewComponentBuilder1.ObjectForNewComponent.Clear()

						'Non Wavelink add body
                        If Not wavelinkfeature Then
							createNewComponentBuilder1.ObjectForNewComponent.Add(body)
							'lw.WriteLine("   Solid body added successfully.")
						End If

						createNewComponentBuilder1.NewFile = fileNew1

						Dim nXObject2 As NXOpen.NXObject = Nothing
						nXObject2 = createNewComponentBuilder1.Commit()
						lw.WriteLine(" ")
						lw.WriteLine(" - Component created for: " & selectedObjectName)

						Dim bodyToAdd As NXOpen.Body = CType(body, NXOpen.Body)

						If smartsortingfeature Then
							If EasyWeightsortinglogic Then
								Try
									materialName = bodyToAdd.GetStringAttribute("EW_Material")
								Catch exInner As Exception
									materialName = "Not specified"
								End Try

								Try
									bodyWeight = bodyToAdd.GetRealAttribute("EW_Body_Weight")
								Catch exInner As Exception
									bodyWeight = -1
								End Try
							Else
								Try
									materialName = GetMaterialName(bodyToAdd)
								Catch exInner As Exception
									materialName = "Not specified"
								End Try

								Try
									bodyWeight = GetBodyWeight(bodyToAdd)
								Catch exInner As Exception
									bodyWeight = -1
								End Try
							End If

							lw.WriteLine(String.Format("   Material Name:         {0}", materialName))
							lw.WriteLine(String.Format("   Weight:                {0}", bodyWeight.ToString()))
						End If

						Dim newComponent As NXOpen.Assemblies.Component = TryCast(nXObject2, NXOpen.Assemblies.Component)
						Dim newComponentPart As Part = CType(newComponent.Prototype, Part)

						If wavelinkfeature Then
							' Change the work part to the new component's part
                            theSession.Parts.SetWork(newComponentPart)

							' Setup the WaveLinkBuilder in the new component's context
                            Dim waveLinkBuilder As Features.WaveLinkBuilder = newComponentPart.BaseFeatures.CreateWaveLinkBuilder(Nothing)
							waveLinkBuilder.Type = Features.WaveLinkBuilder.Types.BodyLink
							Dim extractFaceBuilder As Features.ExtractFaceBuilder = waveLinkBuilder.ExtractFaceBuilder
							extractFaceBuilder.FaceOption = Features.ExtractFaceBuilder.FaceOptionType.FaceChain
							extractFaceBuilder.ParentPart = Features.ExtractFaceBuilder.ParentPartType.OtherPart
							extractFaceBuilder.Associative = wlAssociative
							extractFaceBuilder.FixAtCurrentTimestamp = wlFixAtCurrentTimestamp
							extractFaceBuilder.HideOriginal = wlHideOriginal
							extractFaceBuilder.InheritDisplayProperties = wlInheritDisplayProperties
							extractFaceBuilder.MakePositionIndependent = wlMakePositionIndependent
							extractFaceBuilder.CopyThreads = wlCopyThreads
							Dim selectObjectList As SelectObjectList = extractFaceBuilder.BodyToExtract
							selectObjectList.Add(body)
							waveLinkBuilder.Commit()
							lw.WriteLine("   WaveLink added successfully.")
							waveLinkBuilder.Destroy()
						End If

						If setcomponentflag Then
							SetComponentCreated(body, True)
						End If

						theSession.Parts.SetWork(workPart)

					Catch ex As Exception
						lw.WriteLine(" ")
						lw.WriteLine("Yo ho, mates, we've hit a snag... an error has marooned us: " & ex.Message)
						lw.WriteLine("Pirate's Proclamation: " & ex.StackTrace)
					Finally
						If createNewComponentBuilder1 IsNot Nothing Then
							createNewComponentBuilder1.Destroy()
						End If
						If attributePropertiesBuilder1 IsNot Nothing Then
							attributePropertiesBuilder1.Destroy()
						End If
					End Try
				Else
					Try
						Dim markId2 As NXOpen.Session.UndoMarkId
						markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Component Creator")
						AssemblyidString = assemblyid & tcsecondround
						Dim fileNew1 As NXOpen.FileNew = theSession.Parts.FileNew()
						Dim partOperationCreateBuilder1 As NXOpen.PDM.PartOperationCreateBuilder = Nothing
						partOperationCreateBuilder1 = theSession.PdmSession.CreateCreateOperationBuilder(NXOpen.PDM.PartOperationBuilder.OperationType.Create)
						fileNew1.SetPartOperationCreateBuilder(partOperationCreateBuilder1)
						partOperationCreateBuilder1.SetOperationSubType(NXOpen.PDM.PartOperationCreateBuilder.OperationSubType.FromTemplate)
						partOperationCreateBuilder1.SetModelType("master")
						partOperationCreateBuilder1.SetItemType("Item")
						partOperationCreateBuilder1.CreateLogicalObjects(logicalobjects1)
						sourceobjects1 = logicalobjects1(0).GetUserAttributeSourceObjects()

						partOperationCreateBuilder1.DefaultDestinationFolder = tcDefaultDestinationFolder
						fileNew1.TemplateFileName = tcTemplateFileName
						fileNew1.Units = tcUnits
						fileNew1.RelationType = tcRelationType
						fileNew1.TemplatePresentationName = tcTemplatePresentationName
						fileNew1.ItemType = tcItemType

						fileNew1.UseBlankTemplate = False
						fileNew1.ApplicationName = "ModelTemplate"
						fileNew1.UsesMasterModel = "No"
						fileNew1.TemplateType = NXOpen.FileNewTemplateType.Item
						fileNew1.Specialization = ""
						fileNew1.SetCanCreateAltrep(False)
						partOperationCreateBuilder1.SetAddMaster(False)
						partOperationCreateBuilder1.CreateLogicalObjects(logicalobjects2)
						partOperationCreateBuilder1.SetAddMaster(False)
						Dim nullNXOpen_BasePart As NXOpen.BasePart = Nothing
						Dim objects1(-1) As NXOpen.NXObject
						attributePropertiesBuilder1 = theSession.AttributeManager.CreateAttributePropertiesBuilder(nullNXOpen_BasePart, objects1, NXOpen.AttributePropertiesBuilder.OperationType.Create)
						Dim objects2(-1) As NXOpen.NXObject
						attributePropertiesBuilder1.SetAttributeObjects(objects2)
						Dim objects3(0) As NXOpen.NXObject
						objects3(0) = sourceobjects1(0)
						attributePropertiesBuilder1.SetAttributeObjects(objects3)
						attributePropertiesBuilder1.Title = "DB_PART_NO"
						attributePropertiesBuilder1.Category = "Item"
						attributePropertiesBuilder1.StringValue = AssemblyidString
						attributePropertiesBuilder1.Category = "Item"
						Dim changed1 As Boolean = Nothing
						changed1 = attributePropertiesBuilder1.CreateAttribute()
						Dim attributetitles1(-1) As String
						Dim titlepatterns1(-1) As String
						nXObject1 = partOperationCreateBuilder1.CreateAttributeTitleToNamingPatternMap(attributetitles1, titlepatterns1)
						Dim objects4(0) As NXOpen.NXObject
						objects4(0) = logicalobjects1(0)
						Dim properties1(0) As NXOpen.NXObject
						properties1(0) = nXObject1
						Dim errorList1 As NXOpen.ErrorList = Nothing
						errorList1 = partOperationCreateBuilder1.AutoAssignAttributesWithNamingPattern(objects4, properties1)
						errorList1.Dispose()
						attributePropertiesBuilder1.Title = "DB_PART_NAME"
						attributePropertiesBuilder1.StringValue = selectedObjectName
						attributePropertiesBuilder1.Category = "Item"
						Dim changed2 As Boolean = Nothing
						changed2 = attributePropertiesBuilder1.CreateAttribute()
						fileNew1.MasterFileName = ""
						fileNew1.MakeDisplayedPart = False
						fileNew1.DisplayPartOption = NXOpen.DisplayPartOption.AllowAdditional
						partOperationCreateBuilder1.ValidateLogicalObjectsToCommit()
						Dim logicalobjects4(0) As NXOpen.PDM.LogicalObject
						logicalobjects4(0) = logicalobjects1(0)
						partOperationCreateBuilder1.CreateSpecificationsForLogicalObjects(logicalobjects4)
						createNewComponentBuilder1 = workPart.AssemblyManager.CreateNewComponentBuilder()
						createNewComponentBuilder1.ReferenceSetName = "MODEL"
						createNewComponentBuilder1.ComponentOrigin = NXOpen.Assemblies.CreateNewComponentBuilder.ComponentOriginType.Absolute
						createNewComponentBuilder1.OriginalObjectsDeleted = False
						createNewComponentBuilder1.ObjectForNewComponent.Clear()

						'Non Wavelink add body
                        If Not wavelinkfeature Then
							createNewComponentBuilder1.ObjectForNewComponent.Add(body)
							'lw.WriteLine("   Solid body added successfully.")
						End If

						createNewComponentBuilder1.NewFile = fileNew1

						Dim nXObject2 As NXOpen.NXObject = Nothing
						nXObject2 = createNewComponentBuilder1.Commit()
						lw.WriteLine(" ")
						lw.WriteLine(" - Component created for: " & selectedObjectName)

						Dim bodyToAdd As NXOpen.Body = CType(body, NXOpen.Body)

						If smartsortingfeature Then
							If EasyWeightsortinglogic Then
								Try
									materialName = bodyToAdd.GetStringAttribute("EW_Material")
								Catch exInner As Exception
									materialName = "Not specified"
								End Try

								Try
									bodyWeight = bodyToAdd.GetRealAttribute("EW_Body_Weight")
								Catch exInner As Exception
									bodyWeight = -1
								End Try
							Else
								Try
									materialName = GetMaterialName(bodyToAdd)
								Catch exInner As Exception
									materialName = "Not specified"
								End Try

								Try
									bodyWeight = GetBodyWeight(bodyToAdd)
								Catch exInner As Exception
									bodyWeight = -1
								End Try
							End If

							lw.WriteLine(String.Format("   Material Name:         {0}", materialName))
							lw.WriteLine(String.Format("   Weight:                {0}", bodyWeight.ToString()))
						End If

						Dim newComponent As NXOpen.Assemblies.Component = TryCast(nXObject2, NXOpen.Assemblies.Component)
						Dim newComponentPart As Part = CType(newComponent.Prototype, Part)

						If wavelinkfeature Then
							' Change the work part to the new component's part
                            theSession.Parts.SetWork(newComponentPart)

							' Setup the WaveLinkBuilder in the new component's context
                            Dim waveLinkBuilder As Features.WaveLinkBuilder = newComponentPart.BaseFeatures.CreateWaveLinkBuilder(Nothing)
							waveLinkBuilder.Type = Features.WaveLinkBuilder.Types.BodyLink
							Dim extractFaceBuilder As Features.ExtractFaceBuilder = waveLinkBuilder.ExtractFaceBuilder
							extractFaceBuilder.FaceOption = Features.ExtractFaceBuilder.FaceOptionType.FaceChain
							extractFaceBuilder.ParentPart = Features.ExtractFaceBuilder.ParentPartType.OtherPart
							extractFaceBuilder.Associative = wlAssociative
							extractFaceBuilder.FixAtCurrentTimestamp = wlFixAtCurrentTimestamp
							extractFaceBuilder.HideOriginal = wlHideOriginal
							extractFaceBuilder.InheritDisplayProperties = wlInheritDisplayProperties
							extractFaceBuilder.MakePositionIndependent = wlMakePositionIndependent
							extractFaceBuilder.CopyThreads = wlCopyThreads
							Dim selectObjectList As SelectObjectList = extractFaceBuilder.BodyToExtract
							selectObjectList.Add(body)
							waveLinkBuilder.Commit()
							lw.WriteLine("   WaveLink added successfully.")
							waveLinkBuilder.Destroy()
						End If

						' Mark the original body as processed
                        If setcomponentflag Then
							SetComponentCreated(body, True)
						End If

						theSession.Parts.SetWork(workPart)

					Catch ex As Exception
						lw.WriteLine(" ")
						lw.WriteLine("Yo ho, mates, we've hit a snag... an error has marooned us: " & ex.Message)
						lw.WriteLine("Pirate's Proclamation: " & ex.StackTrace)
					Finally
						If createNewComponentBuilder1 IsNot Nothing Then
							createNewComponentBuilder1.Destroy()
						End If
						If attributePropertiesBuilder1 IsNot Nothing Then
							attributePropertiesBuilder1.Destroy()
						End If
					End Try
				End If
			Else
				' Setup for local (non-Teamcenter) environment
				Try
					Dim markId2 As NXOpen.Session.UndoMarkId
					markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Component Creator")

					Dim fileNew1 As NXOpen.FileNew = theSession.Parts.FileNew()

					' Construct the new file name with the next available ID
					Dim newFileName As String = lldirectoryPath & baseAssemblyId & llnextAvailableId.ToString("D3") & "-" & selectedObjectName & ".prt"
					Dim simpleFileName As String = baseAssemblyId & llnextAvailableId.ToString("D3") & "-" & selectedObjectName & ".prt"

					fileNew1.NewFileName = newFileName

					fileNew1.UseBlankTemplate = False
					fileNew1.ApplicationName = "ModelTemplate"
					fileNew1.Units = llUnits
					fileNew1.TemplateType = NXOpen.FileNewTemplateType.Item
					fileNew1.TemplatePresentationName = llTemplatePresentationName
					fileNew1.AllowTemplatePostPartCreationAction(False)
					fileNew1.TemplateFileName = llTemplateFileName
					fileNew1.MakeDisplayedPart = False

					createNewComponentBuilder1 = workPart.AssemblyManager.CreateNewComponentBuilder()
					createNewComponentBuilder1.DefiningObjectsAdded = False
					createNewComponentBuilder1.NewComponentName = selectedObjectName.ToString
					createNewComponentBuilder1.ReferenceSetName = "MODEL"
					createNewComponentBuilder1.OriginalObjectsDeleted = False
					createNewComponentBuilder1.DefiningObjectsAdded = True
					createNewComponentBuilder1.ComponentOrigin = NXOpen.Assemblies.CreateNewComponentBuilder.ComponentOriginType.Absolute
					createNewComponentBuilder1.ObjectForNewComponent.Clear()
					createNewComponentBuilder1.NewFile = fileNew1
					Dim bodyToAdd As NXOpen.Body = CType(tempComp, NXOpen.Body)
					lw.WriteLine(" ")
					lw.WriteLine(String.Format(" - Processing Body:      " & selectedObjectName))

					If smartsortingfeature Then
						If EasyWeightsortinglogic Then
							Try
								materialName = bodyToAdd.GetStringAttribute("EW_Material")
							Catch exInner As Exception
								materialName = "Not specified"
							End Try

							Try
								bodyWeight = bodyToAdd.GetRealAttribute("EW_Body_Weight")
							Catch exInner As Exception
								bodyWeight = -1 ' Use -1 or another indicative value to signify that the attribute was not found
							End Try
						Else
							Try
								materialName = GetMaterialName(bodyToAdd)
							Catch exInner As Exception
								materialName = "Not specified"
							End Try

							Try
								bodyWeight = GetBodyWeight(bodyToAdd)
							Catch exInner As Exception
								bodyWeight = -1 ' Use -1 or another indicative value to signify that the attribute was not found
							End Try
						End If
						lw.WriteLine(String.Format("   Material Name:        {0}", materialName))
						lw.WriteLine(String.Format("   Weight:               {0}", bodyWeight.ToString()))
					End If

					' Add a selected solid body to the component without Wavelink
					If Not wavelinkfeature Then
						Dim added1 As Boolean = createNewComponentBuilder1.ObjectForNewComponent.Add(bodyToAdd)
						lw.WriteLine("   Solid body added successfully.")
					End If

					nXObject1 = createNewComponentBuilder1.Commit()
					lw.WriteLine("   Component created as: " & simpleFileName)

					If wavelinkfeature Then
						Dim newComponent As NXOpen.Assemblies.Component = CType(nXObject1, NXOpen.Assemblies.Component)
						Dim newComponentPart As NXOpen.Part = CType(newComponent.Prototype, NXOpen.Part)

						' Switch to the new component part to work within its context
						Dim partLoadStatus As NXOpen.PartLoadStatus = Nothing
						theSession.Parts.SetWorkComponent(newComponent, NXOpen.PartCollection.RefsetOption.Current, NXOpen.PartCollection.WorkComponentOption.Visible, partLoadStatus)
						If partLoadStatus IsNot Nothing Then partLoadStatus.Dispose()

						' Setup the WaveLinkBuilder in the new component's context
						Dim waveLinkBuilder As Features.WaveLinkBuilder = newComponentPart.BaseFeatures.CreateWaveLinkBuilder(Nothing)
						waveLinkBuilder.Type = Features.WaveLinkBuilder.Types.BodyLink
						Dim extractFaceBuilder As Features.ExtractFaceBuilder = waveLinkBuilder.ExtractFaceBuilder
						extractFaceBuilder.FaceOption = Features.ExtractFaceBuilder.FaceOptionType.FaceChain
						extractFaceBuilder.ParentPart = Features.ExtractFaceBuilder.ParentPartType.OtherPart
						extractFaceBuilder.Associative = wlAssociative
						extractFaceBuilder.FixAtCurrentTimestamp = wlFixAtCurrentTimestamp
						extractFaceBuilder.HideOriginal = wlHideOriginal
						extractFaceBuilder.InheritDisplayProperties = wlInheritDisplayProperties
						extractFaceBuilder.MakePositionIndependent = wlMakePositionIndependent
						extractFaceBuilder.CopyThreads = wlCopyThreads
						Dim selectObjectList As SelectObjectList = extractFaceBuilder.BodyToExtract

						' Setting up ScCollector and SelectionIntentRule for the body
						Dim scCollector As NXOpen.ScCollector = extractFaceBuilder.ExtractBodyCollector
						Dim selectionIntentRuleOptions As NXOpen.SelectionIntentRuleOptions = newComponentPart.ScRuleFactory.CreateRuleOptions()
						selectionIntentRuleOptions.SetSelectedFromInactive(False)

						Dim bodies() As Body = {bodyToAdd}
						Dim bodyDumbRule As NXOpen.BodyDumbRule = newComponentPart.ScRuleFactory.CreateRuleBodyDumb(bodies, True, selectionIntentRuleOptions)
						selectionIntentRuleOptions.Dispose()

						Dim rules() As NXOpen.SelectionIntentRule = {bodyDumbRule}
						scCollector.ReplaceRules(rules, False)

						waveLinkBuilder.Commit()
						lw.WriteLine("   WaveLink added successfully.")
						waveLinkBuilder.Destroy()
					End If

					' Mark the original body as processed
					If setcomponentflag Then
						SetComponentCreated(body, True)
					End If

					createNewComponentBuilder1.Destroy()
					theSession.CleanUpFacetedFacesAndEdges()
					theSession.Parts.SetWork(workPart)

					' Add the new ID to the set to track it within the session
					usedIds.Add(llnextAvailableId)

					' Find the next available ID based on the fillTheGap setting
					If fillTheGap Then
						llnextAvailableId += 1
						While usedIds.Contains(llnextAvailableId)
							llnextAvailableId += 1
						End While
					Else
						llnextAvailableId = usedIds.Max + 1
					End If

				Catch ex As NXOpen.NXException When ex.Message.Contains("File already exists")
					lw.WriteLine(" ")
					lw.WriteLine("We attempted to fill the gap during component creation, but")
					lw.WriteLine("encountered an error because one or more removed parts are still")
					lw.WriteLine("in memory. Please close them in the NX session as well.")
					lw.WriteLine("Go to File > Close > Selected Parts.")
				Catch ex As Exception
					lw.WriteLine("By Blackbeard's ghost, we're in uncharted waters... a complication has arisen: " & ex.Message)
				Finally
				End Try
			End If
		Next
		lw.WriteLine(" ")

		Dim endQuotes As New List(Of String) From {
			"Our expedition into the dusk reaches its twilight. Now, who recalls the spot of our anchorage?",
			"Our odyssey across the realms of power concludes.",
			"Our dance with destiny ends in silence.",
			"Our voyage through the storm finds its harbor in the void.",
			"Our voyage has sailed into the sunset. Now, who remembers where we parked?",
			"We've run out of road. Next stop: uncharted couch territories.",
			"That's a wrap on our adventure. Please exit through the gift shop.",
			"The end of our quest is here. Time to hang up our capes.",
			"We've navigated the void and returned. Yet, the darkness lingers, an eternal companion.",
			"Our expedition's final log. Beam us up, there's no intelligent life down here!",
			"Our shared path diverges here. May your socks always match in future adventures.",
			"The torch of our adventure dims, its light flickering one final moment."
		}

		Dim rnd As New Random()
		Dim index As Integer = rnd.Next(endQuotes.Count)
		Dim selectedQuote As String = endQuotes(index)

		lw.WriteLine(selectedQuote)
		lw.WriteLine(" ")
		'lw.WriteLine("----------")
	End Sub

	Function IsComponentCreated(ByVal body As Body) As Boolean
		Try
			Dim attrValue As String = ""

			' Determine the target body based on whether it's an occurrence
			Dim targetBody As Body = If(body.IsOccurrence, body.Prototype, body)

			' Check if the attribute exists and retrieve its value
			If targetBody.HasUserAttribute("Component_created", NXObject.AttributeType.String, -1) Then
				attrValue = targetBody.GetStringAttribute("Component_created")
			End If

			' If the attribute exists but is empty, interpret it as "no" (False)
			If String.IsNullOrEmpty(attrValue) Then
				Return False
			End If
			
			' If the attribute value is a valid boolean string, return its boolean equivalent
			If attrValue.Equals("True", StringComparison.OrdinalIgnoreCase) OrElse
			   attrValue.Equals("False", StringComparison.OrdinalIgnoreCase) Then
				Return Boolean.Parse(attrValue)
			Else
				' If the attribute value is not a recognized boolean string, log a message and interpret as False
				lw.WriteLine("Arr, this 'Component_created' be flying a foreign flag: '" & attrValue & "' for " & If(body.IsOccurrence, "instance: ", "body: ") & targetBody.JournalIdentifier)
				Return False
			End If

		Catch ex As NXOpen.NXException
			lw.WriteLine("Shiver me timbers, we've sailed into a storm... a mistake has been spotted: " & ex.Message)
		End Try

		' Return false if attribute not found, not valid, or any exception occurs
		Return False
	End Function

	Sub SetComponentCreated(ByVal body As Body, ByVal created As Boolean)
		Try
			Dim targetBody As Body = body

			' If the body is an occurrence, use the prototype body for setting attributes
			If body.IsOccurrence Then
				targetBody = body.Prototype
			End If

			' Set the user attribute on the target body
			targetBody.SetUserAttribute("Component_created", -1, created.ToString(), Update.Option.Now)

			' Log a success message indicating the attribute was set
			'lw.WriteLine("Attribute 'Component_created' set to " & created.ToString() & " for " & If(body.IsOccurrence, "instance: ", "body: ") & targetBody.JournalIdentifier)

		Catch ex As NXOpen.NXException
			lw.WriteLine("Hitch in casting the line 'Component_created' on " & If(body.IsOccurrence, "instance: ", "body: ") & body.JournalIdentifier & " - " & ex.Message)
		End Try
	End Sub

	Function SelectObjects(prompt As String, ByRef dispObj As List(Of DisplayableObject)) As Boolean
		Dim selObj As NXObject()
		Dim title As String = "Select solid bodies"
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

		If resp = Selection.Response.ObjectSelected OrElse
		   resp = Selection.Response.ObjectSelectedByName OrElse
		   resp = Selection.Response.Ok Then

			If selObj IsNot Nothing AndAlso selObj.Length > 0 Then
				For Each item As NXObject In selObj
					If String.IsNullOrEmpty(item.Name) Then
						item.SetName(defaultsolidbodyname)
					End If
					dispObj.Add(CType(item, DisplayableObject))
				Next

				' Update EW_Body_Weight if EasyWeightsortinglogic is true
				If EasyWeightsortinglogic Then
					For Each body As DisplayableObject In dispObj
						If TypeOf body Is Body Then
							UpdateBodyWeight(CType(body, Body))
						End If
					Next
					lw.WriteLine("")
					lw.WriteLine(" - Successfully updated the Weight information.")
					lw.WriteLine("")
				End If

				' SmartSort objects
				If smartsortingfeature Then
					If EasyWeightsortinglogic Then
						Try
							dispObj.Sort(Function(a, b)
											 Dim aMat As String = Nothing
											 Dim bMat As String = Nothing
											 Dim aWeight As Double = 0
											 Dim bWeight As Double = 0
											 Dim primaryResult As Integer = 0

											 Try
												 aMat = a.GetStringAttribute("EW_Material")
											 Catch ex As Exception
												 aMat = "zzzzz"
											 End Try

											 Try
												 bMat = b.GetStringAttribute("EW_Material")
											 Catch ex As Exception
												 bMat = "zzzzz"
											 End Try

											 Try
												 aWeight = a.GetRealAttribute("EW_Body_Weight")
											 Catch ex As Exception
												 aWeight = 0
											 End Try

											 Try
												 bWeight = b.GetRealAttribute("EW_Body_Weight")
											 Catch ex As Exception
												 bWeight = 0
											 End Try

											 Dim aNum As Double? = GetMaterialThickness(aMat)
											 Dim bNum As Double? = GetMaterialThickness(bMat)

											 ' Handling primary sort based on EW_Material attribute
											 If aNum.HasValue And bNum.HasValue Then
												 primaryResult = bNum.Value.CompareTo(aNum.Value) ' Sort in descending order
											 ElseIf aNum.HasValue Then
												 primaryResult = -1
											 ElseIf bNum.HasValue Then
												 primaryResult = 1
											 Else
												 primaryResult = String.Compare(aMat, bMat) ' Sort alphabetically in that case
											 End If

											 ' Handling secondary sort based on EW_Body_Weight attribute
											 If primaryResult = 0 Then
												 Return bWeight.CompareTo(aWeight) ' Sort in descending order based on weight
											 Else
												 Return primaryResult ' Otherwise, return the result of the primary comparison
											 End If
										 End Function)
						Catch ex As Exception
							lw.WriteLine(" ")
							lw.WriteLine("Hoist the colors, we're navigating choppy seas... a fault has been discovered: " & ex.Message)
						End Try
					Else
						Try
							dispObj.Sort(Function(a, b)
											 Dim aMat As String = If(GetMaterialName(a) = "Not specified", "zzzzz", GetMaterialName(a))
											 Dim bMat As String = If(GetMaterialName(b) = "Not specified", "zzzzz", GetMaterialName(b))

											 Dim aNumVal As Double? = GetMaterialThickness(aMat)
											 Dim bNumVal As Double? = GetMaterialThickness(bMat)

											 ' Compare numerical values if both are present
											 If aNumVal.HasValue AndAlso bNumVal.HasValue Then
												 Dim numCompare As Integer = bNumVal.Value.CompareTo(aNumVal.Value)
												 If numCompare <> 0 Then Return numCompare
											 ElseIf aNumVal.HasValue Then
												 Return -1
											 ElseIf bNumVal.HasValue Then
												 Return 1
											 End If

											 ' If numerical values are equal or not present, compare the rest of the material name
											 Dim restCompare As Integer = String.Compare(aMat, bMat)
											 If restCompare <> 0 Then Return restCompare

											 ' If materials are identical, compare weights
											 Dim aWeight As Double = GetBodyWeight(a)
											 Dim bWeight As Double = GetBodyWeight(b)
											 Return bWeight.CompareTo(aWeight) ' Sort by weight in descending order
										 End Function)
						Catch ex As Exception
							lw.WriteLine(" ")
							lw.WriteLine("Ahoy, deckhands, a squall's upon us... an anomaly has presented itself: " & ex.Message)
						End Try
					End If

					lw.WriteLine(" - Selected bodies captured and the Selection order: Sorted.")
				Else
					lw.WriteLine(" - Selected bodies captured and the Selection order: Preserved.")
				End If

				Return True ' Successfully selected and sorted objects
			Else
				' Handle the case where no objects are selected
				lw.WriteLine("The chronicle paused, as no items were marked for the journey.")
				Return False
			End If
		Else
			lw.WriteLine(" ")
			lw.WriteLine("Arr, what's this? A baffling response during the selection of the bounty: " & resp.ToString())
			Return False
		End If
	End Function

	Sub UpdateBodyWeight(ByVal body As Body)
		Dim myMeasure As MeasureManager = workPart.MeasureManager
		Dim massUnits(1) As Unit
		massUnits(0) = workPart.UnitCollection.GetBase("Volume")

		Dim mb As MeasureBodies = myMeasure.NewMassProperties(massUnits, 0.99, New Body() {body})

		' Update the InformationUnit for MeasureBodies based on unit system
		Dim informationUnit As MeasureBodies.AnalysisUnit
		If unitString = "in" Then
			mb.informationUnit = MeasureBodies.AnalysisUnit.PoundInch
		Else
			mb.informationUnit = MeasureBodies.AnalysisUnit.KilogramMilliMeter
		End If

		' Extract volume
		Dim bodyVolume As Double = mb.Volume
		mb.Dispose()

		' Extract density from the EW_Material_Density attribute; default to 1 if not found
		Dim density As Double = 1.0
		Try
			density = Convert.ToDouble(body.GetStringAttribute("EW_Material_Density"))
		Catch ex As Exception
			' If the attribute is not found or cannot be converted, use the default density of 1
			'lw.WriteLine("Density attribute not found or invalid for body: " & body.JournalIdentifier & ". Using default density of 1.")
		End Try

		If unitString = "in" Then
			' Calculate weight assuming density is in Pound/Cubic Foot, converting to lbm
			Dim bodyWeight As Double = bodyVolume / 1728 * density
			Try
				body.SetUserAttribute("EW_Body_Weight", -1, bodyWeight, Update.Option.Now)
				'lw.WriteLine("Updated EW_Body_Weight for: " & body.JournalIdentifier & " to " & bodyWeight.ToString("F3") & " Lbm")
			Catch ex As Exception
				'lw.WriteLine("Failed to update EW_Body_Weight for: " & body.JournalIdentifier & ". Error: " & ex.Message)
			End Try
		Else
			' Calculate weight assuming density is in Kg/Cubic Meter, converting to kg
			Dim bodyWeight As Double = bodyVolume / 1000000000.0 * density
			Try
				body.SetUserAttribute("EW_Body_Weight", -1, bodyWeight, Update.Option.Now)
				'lw.WriteLine("Updated EW_Body_Weight for body: " & body.JournalIdentifier & " to " & bodyWeight.ToString("F3") & " Kg")
			Catch ex As Exception
				'lw.WriteLine("Failed to update EW_Body_Weight for body: " & body.JournalIdentifier & ". Error: " & ex.Message)
			End Try
		End If
	End Sub

	Function GetMaterialName(body As Body) As String
		' Retrieve the material name for the body
		Dim matName As String = ""
		Try
			matName = body.GetStringAttribute("Material")
		Catch ex As Exception
			Return If(matName Is Nothing, matName, "Not specified")
		End Try
		Return matName
	End Function

	Function GetMaterialThickness(materialName As String) As Double?
		' Try to extract numerical value
		Dim pattern As String = "(\d+/\d+)|(\d+(\.\d+)?)"
		Dim matches As MatchCollection
		Dim thickness As Double? = Nothing
		Dim numericPart As String

		If materialName.Contains(ssunitmm) Then
			numericPart = materialName.Substring(0, materialName.IndexOf(ssunitmm)).Trim()
			matches = Regex.Matches(numericPart, pattern)
			'lw.WriteLine("Material name trimed (" & ssunitmm & ") : " & numericPart.ToString())
		ElseIf materialName.Contains(ssunitin) Then
			numericPart = materialName.Substring(0, materialName.IndexOf(ssunitin)).Trim()
			matches = Regex.Matches(numericPart, pattern)
			'lw.WriteLine("Material name trimed (" & ssunitin & ") : " & numericPart.ToString())
		Else
			matches = Regex.Matches(materialName, pattern)
		End If

		For Each match As Match In matches
			If match.Success Then
				Dim value As Double
				If match.Value.Contains("/") Then
					Dim parts As String() = match.Value.Split("/")
					If parts.Length = 2 Then
						Dim numerator As Double
						Dim denominator As Double
						If Double.TryParse(parts(0), numerator) AndAlso Double.TryParse(parts(1), denominator) AndAlso denominator <> 0 Then
							thickness = (numerator / denominator)
							If materialName.Contains(ssunitin) Then
								thickness *= 25.4
							End If
							'lw.WriteLine("Calculated Thickness: " & thickness.ToString() & " - from (Fraction)")
							Return thickness
						End If
						Return thickness
					End If
				ElseIf Double.TryParse(match.Value, value) Then
					thickness = value
					If materialName.Contains(ssunitin) Then
						thickness *= 25.4
					End If
					'lw.WriteLine("Calculated Thickness: " & thickness.ToString() & " - from whole or decimal number")
					Return thickness
				End If
			End If
		Next
		Return Nothing
	End Function

	Function GetBodyWeight(body As Body) As Double
		Dim weight As Double = 0.0
		Dim myMeasure As MeasureManager = theSession.Parts.Work.MeasureManager
		Dim theBodies(0) As Body
		theBodies(0) = body

		Dim massUnits(4) As NXOpen.Unit
		massUnits(2) = theSession.Parts.Work.UnitCollection.GetBase("Mass")

		Try
			Dim mb As MeasureBodies = myMeasure.NewMassProperties(massUnits, 0.99, theBodies)
			weight = mb.Mass
			mb.Dispose()
		Catch ex As Exception
			lw.WriteLine(" ")
			lw.WriteLine("Arr, we've hit a snag in weighing our cargo: " & ex.Message)
		End Try
		Return weight
	End Function
End Module