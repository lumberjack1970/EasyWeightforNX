' EasyWeight
' Journal desciption: Alters the color of selected faces. Has priority over the main Material Journal. Used to distinguish the inside/outside of the body.
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13
' V100 - Initial Release - November 2023
' V101 - Configuration Settings

Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF

Module NXJournal
    Dim theSession As Session = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()
    Dim theUI As UI = UI.GetUI()
    Dim lw As ListingWindow = theSession.ListingWindow


    '------------------------
    ' Configuration Settings:

    ' Face Settings:
    Dim facecolor As Double = 17 ' Set the face color to ID: 17
    Dim facename As String = "Inside"
    '------------------------


    Sub Main()
        Dim markId1 As Session.UndoMarkId
        markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Face Color")
        Dim selobj As NXObject
        Dim type As Integer
        Dim subtype As Integer
        Dim theFaces As New List(Of Face)
        Dim theUI As UI = UI.GetUI
        Dim numsel As Integer = theUI.SelectionManager.GetNumSelectedObjects()

        ' Process the preselected Faces
        If numsel > 0 Then
            For inx As Integer = 0 To numsel - 1
                selobj = theUI.SelectionManager.GetSelectedTaggedObject(inx)
                theUfSession.Obj.AskTypeAndSubtype(selobj.Tag, type, subtype)
                If type = UFConstants.UF_solid_type Then
                    theFaces.Add(selobj)
                End If
            Next
        Else
            ' Prompt to select Faces
            If SelectFaces("Select Faces to change color", theFaces) = Selection.Response.Cancel Then
                Return
            End If
        End If

        For Each temp As Face In theFaces
            Dim displayModification As DisplayModification = theSession.DisplayManager.NewDisplayModification()
            With displayModification
                ' .ApplyToAllFaces = False
                .ApplyToOwningParts = True
                .NewColor = facecolor
                .Apply({temp})
            End With
            displayModification.Dispose()
        Next
    End Sub

    Function SelectFaces(ByVal prompt As String, ByRef selFace As List(Of Face)) As Selection.Response
        Dim theUI As UI = UI.GetUI
        Dim title As String = ("Select Faces - " & facename)
        Dim includeFeatures As Boolean = False
        Dim keepHighlighted As Boolean = False
        Dim selAction As Selection.SelectionAction = Selection.SelectionAction.ClearAndEnableSpecific
        Dim scope As Selection.SelectionScope = Selection.SelectionScope.WorkPart
        Dim selectionMask_array(0) As Selection.MaskTriple
        Dim selObj() As TaggedObject

        With selectionMask_array(0)
            .Type = UFConstants.UF_solid_type
            .Subtype = 0
            .SolidBodySubtype = UFConstants.UF_UI_SEL_FEATURE_ANY_FACE
        End With

        Dim resp As Selection.Response = theUI.SelectionManager.SelectTaggedObjects(prompt,
        title, scope, selAction,
        includeFeatures, keepHighlighted, selectionMask_array, selObj)

        If resp = Selection.Response.Ok Then
            For Each temp As TaggedObject In selObj
                selFace.Add(temp)
            Next
            Return Selection.Response.Ok
        Else
            Return Selection.Response.Cancel
        End If
    End Function

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
    End Function

End Module
