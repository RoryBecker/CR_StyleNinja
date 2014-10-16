Imports System.ComponentModel
Imports DevExpress.CodeRush.StructuralParser
Imports SP = DevExpress.CodeRush.StructuralParser
Imports DevExpress.CodeRush.Core
Imports System.Diagnostics.CodeAnalysis
Namespace SA14XX
    Friend Module Rules
#Region "Utility"
        Private MemberOrType As LanguageElementType() = New LanguageElementType() {LanguageElementType.Method, _
                                                                                   LanguageElementType.Class, _
                                                                                   LanguageElementType.Struct, _
                                                                                   LanguageElementType.Interface}
        Friend Function IsMemberOrType(ByVal CodeActive As LanguageElement) As Boolean
            Return MemberOrType.Contains(CodeActive.ElementType)
        End Function
        Friend Function HasExplicitVisibility(ByVal Element As IElement) As Boolean
            Dim ASE = TryCast(Element, AccessSpecifiedElement)
            If ASE Is Nothing Then
                Return False
            End If
            Return ASE.VisibilityRange <> SourceRange.Empty
        End Function
#End Region

#Region "SA1400 + Fix"
        Public Const Message_SA1400 As String = "SA1400 - Access Modifier Must Be Declared"
        Public Sub SA1400_Available(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1400(ea.CodeActive)
        End Sub
        Public Function Qualifies_SA1400(ByVal Element As IElement) As Boolean
            If TypeOf Element Is EnumElement Then
                ' Work around DXCore bug which thinks EnumElement is an AccessSpecifiedElement
                Return False
            End If
            Return ShouldHaveExplicitVisibility(Element) _
                   AndAlso Not HasExplicitVisibility(Element)
        End Function
        Private Function ShouldHaveExplicitVisibility(ByVal Element As IElement) As Boolean
            ' Exclude if not AccessSpecifiedElement
            If Not TypeOf Element Is AccessSpecifiedElement Then
                Return False
            End If
            ' Exclude Params
            If TypeOf Element Is SP.Param Then
                Return False
            End If
            ' Exclude Interface
            If TypeOf Element Is SP.Interface Then
                Return False
            End If
            ' Exclude if Parent is interface
            If Element.Parent IsNot Nothing _
                AndAlso TypeOf Element.Parent Is SP.Interface Then
                Return False
            End If
            ' Exclude if Parent is Method or accessor
            If Element.ParentMethodOrAccessor IsNot Nothing Then
                Return False
            End If
            Return True
        End Function
        Public Sub Fix_SA1400(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            Dim Element = CType(ea.CodeActive, AccessSpecifiedElement)
            ea.TextDocument.InsertText(Element.Range.Start, GetVisibilityName(Element.Visibility) & " ")
        End Sub
#End Region
#Region "SA1401 + Fix"
        Public Const Message_SA1401 As String = "SA1401 - Fields must be Private"
        Public Sub Available_SA1401(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1401(ea.CodeActive)
        End Sub
        Public Function Qualifies_SA1401(ByVal Element As IElement) As Boolean
            Dim Variable = TryCast(Element, Variable)
            If Variable Is Nothing Then
                Return False
            End If
            Return Variable.IsField AndAlso Variable.Visibility <> MemberVisibility.Private
        End Function
        Public Sub Fix_SA1401(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            Dim Element = CType(ea.CodeActive, Variable)
            Dim Visibility = GetVisibilityName(MemberVisibility.Private)
            If Element.VisibilityRange.IsEmpty Then
                ea.TextDocument.InsertText(Element.Range.Start, Visibility & " ")
            Else
                ea.TextDocument.Replace(Element.VisibilityRange, Visibility, "")
            End If
        End Sub

#End Region
#Region "SA1402"
        Public Const Message_SA1402 As String = "SA1402 - Files must only contain a single Class"
        Public Sub SA1402_Available(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1402(ea.CodeActive)
        End Sub
        Public Function Qualifies_SA1402(ByVal Element As IElement) As Boolean
            Dim TheClass = TryCast(Element, [Class])
            If TheClass Is Nothing Then
                Return False
            End If
            Return TheClass.Parent.Nodes.OfType(Of [Class]).Count > 1
        End Function
#End Region
#Region "SA1403"
        Public Const Message_SA1403 As String = "SA1403 - Files must only contain a single Namespace"
        Public Sub SA1403_Available(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1403(ea.CodeActive)
        End Sub
        Public Function Qualifies_SA1403(ByVal Element As IElement) As Boolean
            Dim TheNamespace = TryCast(Element, [Namespace])
            If TheNamespace Is Nothing Then
                Return False
            End If
            Return TheNamespace.FileNode.Nodes.OfType(Of [Namespace]).Count > 1
        End Function
        Public Sub Fix_SA1403(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            ' point at coderush's functionality
            ' or 
            ' cut / paste range elsewhere

        End Sub

#End Region

#Region "SA1404 + No Fix Possible"
        Public Const Message_SA1404 As String = "SA1404 - SuppressMessage attribute does not include a justification"
        Public Sub Available_SA1404(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1404(ea.CodeActive)
        End Sub
        Public Function Qualifies_SA1404(ByVal Element As IElement) As Boolean
            Dim Attribute = TryCast(Element, Attribute)
            If Attribute Is Nothing Then
                Return False
            End If
            Return Attribute.Name = "System.Diagnostics.CodeAnalysis.SuppressMessageAttribute" _
            AndAlso Not Attribute.DetailNodes.OfType(Of AttributeVariableInitializer).Any(Function(a) a.Name = "Justification")
        End Function
        ' No fix possible
#End Region
#Region "SA1405 + No Fix Possible"
        Public Const Message_SA1405 As String = "SA1405 - Debug.Assert does not include a message"
        Public Sub Available_SA1405(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1405(ea.CodeActive)
        End Sub
        Public Function Qualifies_SA1405(ByVal Element As IElement) As Boolean
            Dim MethodCall = TryCast(Element, MethodCall)
            If MethodCall Is Nothing Then
                Return False
            End If
            Return MethodCall.GetDeclaration.FullName = "System.Diagnostics.Debug.Assert" _
            AndAlso MethodCall.DetailNodes.Count = 1 ' No Message
        End Function
        ' No fix possible
#End Region
#Region "SA1406 + Rule can never happen"
        ' Imposible Issue
        'Public Const Message_SA1406 As String = "Debug.Fail() does not specify a descriptive message"
#End Region
#Region "SA1409 + Fix"
        Public Const Message_SA1409 As String = "SA1409 - Remove Unnecessary Code"
        Public Sub Available_SA1409(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1409(ea.CodeActive)
        End Sub
        Public Function Qualifies_SA1409(ByVal Element As IElement) As Boolean
            Dim TheTry = TryCast(Element, [Try])
            If TheTry Is Nothing Then
                Return False
            End If
            Return TheTry.NodeCount = 0 AndAlso TheTry.NextSibling.NodeCount = 0
        End Function
        Public Sub Fix_SA1409(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            ea.TextDocument.DeleteText(ea.CodeActive.GetFullBlockCutRange())
        End Sub
#End Region
    End Module
End Namespace
