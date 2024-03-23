' Written by Tamas Woller - March 2024, V101
' Journal desciption: you can control the visibility of specific solid bodies on your screen using the material names assigned before.
' Shared on NXJournaling.com
' Written in VB.Net
' Tested on Siemens NX 2212 and 2306, Native and Teamcenter 13
' V100 - Initial Release - November 2023
' V101 - Added Configuration settings

Option Strict Off
Imports System
Imports NXOpen
Imports System.Collections.Generic
Imports System.Drawing

Module Module1
    Private _attributeValue As String '= ""

    '------------------------
    ' Configuration Settings:
    ' The first item in the list will appear in the display window, while the second item represents the material name the system searches for within the solid body.

    Public materials = New String() {
        "Plywood - 12mm ,12mm Plywood",
        "Softwood ,SOFTWOOD",
        "Hardwood ,HARDWOOD",
        "Upholstery - Fabric ,Fabric",
        "Upholstery - Cushion ,Cushion",
        "Fiddle ,FIDDLE",
        "Mirror ,MIRROR",
        "Stainless Steel ,Stainless Steel",
        "Without Weight ,empty"
    }
    '------------------------

    Sub UpdateView(ByRef theSession As Session, workPart As Part, attributeValue As String)
        Dim bodies() As Body = workPart.Bodies.ToArray()
        Dim bodiesToShow As New List(Of Body)
        Dim bodiesToHide As New List(Of Body)

        For Each body As Body In bodies
            Dim materialName As String = String.Empty
            Try
                ' Attempt to retrieve the material name from the body's attributes
                materialName = body.GetStringAttribute("Material")

                If attributeValue = "empty" Then
                    bodiesToHide.Add(body)
                ElseIf materialName.Equals(attributeValue, StringComparison.OrdinalIgnoreCase) Then
                    bodiesToShow.Add(body)
                Else
                    bodiesToHide.Add(body)
                End If
            Catch ex As NXException
                If attributeValue = "empty" Then
                    bodiesToShow.Add(body)
                Else
                    bodiesToHide.Add(body)
                End If
            End Try
        Next

        ' Hide bodies not matching the selected material
        If bodiesToHide.Count > 0 Then
            theSession.DisplayManager.BlankObjects(bodiesToHide.ToArray())
        End If

        ' Show bodies matching the selected material
        If bodiesToShow.Count > 0 Then
            theSession.DisplayManager.UnblankObjects(bodiesToShow.ToArray())
        End If
    End Sub

    Sub Main()

        Dim theSession As Session = Session.GetSession()
        Dim markId1 As NXOpen.Session.UndoMarkId
        markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Material Filter")
        If IsNothing(theSession.Parts.Work) Then
            Return
        End If

        Dim workPart As Part = theSession.Parts.Work
        Dim lw As ListingWindow = theSession.ListingWindow
        lw.Open()

        Dim myForm As New Form1
        myForm.AttributeTitle = "Select One from the Following Options:"

        AddHandler myForm.ApplyButtonClicked, Sub(sender, e)
                                                  _attributeValue = myForm.AttributeValue
                                                  UpdateView(theSession, workPart, _attributeValue)

                                                  ' Force NX to refresh its graphics display
                                                  theSession.Parts.Work.Views.Refresh()
                                              End Sub

        Do
            myForm.ShowDialog()
            If myForm.Canceled Then
                Exit Do
            End If
        Loop
        lw.Close()
    End Sub


    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately
    End Function
End Module

Public Class Form1
    Public Event ApplyButtonClicked As EventHandler
    Private materialDictionary As New Dictionary(Of String, String)
    Private _frmAttributeTitle As String
    Public Property AttributeTitle() As String
        Get
            Return _frmAttributeTitle
        End Get
        Set(ByVal value As String)
            _frmAttributeTitle = value
        End Set
    End Property

    Private _frmAttributeValue As String
    Public Property AttributeValue() As String
        Get
            Return _frmAttributeValue
        End Get
        Set(ByVal value As String)
            _frmAttributeValue = value
        End Set
    End Property

    Private _canceled As Boolean = False
    Public ReadOnly Property Canceled() As Boolean
        Get
            Return _canceled
        End Get
    End Property

	Private Sub Form1_Load(sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		Label1.Text = _frmAttributeTitle
		PopulateMaterialList()  ' This line calls the PopulateMaterialList method to fill the ListBox
		ListBox1.SelectedItem = _frmAttributeValue
						
		' Set the color properties
		Me.BackColor = Color.FromArgb(55, 55, 55)
		btnApply.BackColor = Color.FromArgb(50, 50, 50)
		btnQuit.BackColor = Color.FromArgb(50, 50, 50)

		' Set the font colors to white
		btnApply.ForeColor = Color.White
		btnQuit.ForeColor = Color.White
		Label1.ForeColor = Color.White
		
	End Sub		
		
	Private Sub PopulateMaterialList()
		' Clear existing entries to avoid duplicates
		materialDictionary.Clear()
		ListBox1.Items.Clear() ' Clear existing ListBox items

		For Each material As String In materials
			Dim parts = material.Split(",")
			Dim displayName = parts(0).Trim()
			Dim attributeValue = parts(1).Trim()

			' Skip adding if key already exists
			If Not materialDictionary.ContainsKey(displayName) Then
				materialDictionary.Add(displayName, attributeValue)
				ListBox1.Items.Add(displayName)  ' Populate the ListBox
			End If
		Next
	End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
    'MsgBox("Apply button clicked") ' Debugging line to ensure the method is triggered
    If ListBox1.SelectedItem IsNot Nothing Then
        Dim selectedAlias As String = ListBox1.SelectedItem.ToString()
        _frmAttributeValue = materialDictionary(selectedAlias)
        RaiseEvent ApplyButtonClicked(Me, EventArgs.Empty)
    Else
        ' You might want to display a warning to select an item
    End If
	End Sub

    Private Sub btnQuit_Click(sender As Object, e As EventArgs) Handles btnQuit.Click

    Dim theSession As Session = Session.GetSession()
    If Not IsNothing(theSession.Parts.Work) Then
        Dim workPart As Part = theSession.Parts.Work
    End If
    
    _canceled = True
    Me.Close()
	End Sub
End Class
 
 <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer, or just start playing with the numbers
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.btnQuit = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
	    Me.ListBox1.Location = New System.Drawing.Point(220, 20) ' This is the starting point for the ListBox in the Window - Width, Height
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(150, 400) ' This is the Overall size of the ListBox - Width, Height
        Me.ListBox1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(5, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(123, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "SELECT_MATERIAL"
        '
        'btnApply
        '
        Me.btnApply.Location = New System.Drawing.Point(12, 386) ' This is the starting point for the Button in the Window - Width, Height
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(75, 23) ' This is the Overall size of the Button - Width, Height
        Me.btnApply.TabIndex = 3
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'btnQuit
        '
        Me.btnQuit.Location = New System.Drawing.Point(119, 386) ' This is the starting point for the Button in the Window - Width, Height
        Me.btnQuit.Name = "btnQuit"
        Me.btnQuit.Size = New System.Drawing.Size(75, 23) ' This is the Overall size of the Button - Width, Height
        Me.btnQuit.TabIndex = 4
        Me.btnQuit.Text = "Quit"
        Me.btnQuit.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(380, 420) ' This is the Overall size of the Window - Width, Height
        Me.Controls.Add(Me.btnQuit)
        Me.Controls.Add(Me.btnApply)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListBox1)
        Me.Name = "Form1"
        Me.Text = "Solid Body Filter Tool"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnApply As System.Windows.Forms.Button
    Friend WithEvents btnQuit As System.Windows.Forms.Button

End Class