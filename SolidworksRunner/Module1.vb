
Imports System.CodeDom
Imports System.ComponentModel.Design
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst


Module Module1

    Sub Main()
        Dim app As SldWorks = New SldWorks()
        Dim doc As ModelDoc2 = app.ActiveDoc
        Debug.WriteLine(doc.GetTitle())

        Dim materialMap As Dictionary(Of String, Component2) = New Dictionary(Of String, Component2)
        Dim rootComponent As Component2 = doc.ConfigurationManager.ActiveConfiguration().GetRootComponent()
        For Each component As Component2 In getAllSubChildren(rootComponent)
            Debug.WriteLine(component)
        Next

    End Sub





    ' Helper Functions
    Function getAllSubChildren(component As Component2)
        Dim componentsToReturn As List(Of Component2) = New List(Of Component2)
        recurseChildren(component, componentsToReturn)
        Return componentsToReturn
    End Function
    Function recurseChildren(component As Component2, components As List(Of Component2))
        If Not component.GetChildren() Is Nothing Then
            For Each child As Component In component.GetChildren()
                components.Add(child)
                If Not child.GetChildren() Is Nothing Then
                    recurseChildren(child, components)
                End If
            Next
        End If
        Return components
    End Function

End Module
