' EasyWeight
' Journal desciption: Keeps the solid body unchanged but removes any weight-related EasyWeight attributes.
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13
' V100 - Initial Release - November 2023

Imports System
Imports NXOpen
Imports System.Collections.Generic
Imports NXOpen.UF

Module NXJournal
    Dim theSession As Session
    Dim theUFSession As UFSession
    Dim workPart As Part
    Dim displayPart As Part
    Dim mySelectedObjects As New List(Of DisplayableObject)
    Dim theUI As UI

    Sub Main(ByVal args() As String)
        Try
            theSession = Session.GetSession()
            theUFSession = UFSession.GetUFSession()
            workPart = theSession.Parts.Work
            displayPart = theSession.Parts.Display
            theUI = UI.GetUI()

            Dim markId1 As Session.UndoMarkId = Nothing
            markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Delete Attributes")

            Dim lw As ListingWindow = theSession.ListingWindow
            lw.Open()

            If SelectObjects("Select bodies", mySelectedObjects) = Selection.Response.Ok Then
                ' Convert DisplayableObjects to Bodies
                Dim selectedBodies As New List(Of Body)
                For Each obj As DisplayableObject In mySelectedObjects
                    If TypeOf obj Is Body Then
                        selectedBodies.Add(CType(obj, Body))
                    End If
                Next

                ' Delete user attributes
                For Each body As Body In selectedBodies
                    DeleteUserAttribute(body, "EW_Material")
                    DeleteUserAttribute(body, "EW_Body_Weight")
                    DeleteUserAttribute(body, "EW_Material_Density")
                Next

                ' Commit the changes
                theSession.UpdateManager.DoUpdate(markId1)
            Else
                lw.WriteLine("No objects were selected.")
            End If
        Catch ex As Exception
            ' Handle exceptions here
            Console.WriteLine("An error occurred: " & ex.Message)
        End Try
    End Sub

    Function SelectObjects(prompt As String,
                           ByRef dispObj As List(Of DisplayableObject)) As Selection.Response
        Dim selObj As NXObject()
        Dim title As String = "Select solid bodies to DELETE their Weight Data"
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

    Sub DeleteUserAttribute(ByVal theObject As NXObject, ByVal attributeName As String)
        Dim attributeInfo() As NXObject.AttributeInformation = CType(theObject, Body).GetUserAttributes()

        For Each temp As NXObject.AttributeInformation In attributeInfo
            If temp.Title = attributeName Then
                theObject.DeleteUserAttribute(temp.Type, temp.Title, False, Update.Option.Now)
                Exit For
            End If
        Next
    End Sub

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
    End Function
End Module