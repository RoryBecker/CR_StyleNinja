Imports System.ComponentModel
Imports DevExpress.CodeRush.StructuralParser
Imports DevExpress.CodeRush.Core
Imports System.Diagnostics.CodeAnalysis
Imports System.Runtime.CompilerServices

Namespace SA11XX
    Friend Module Rules
#Region "Utility"
        Private MemberOrType As LanguageElementType() = New LanguageElementType() {LanguageElementType.Method, _
                                                                                   LanguageElementType.Class, _
                                                                                   LanguageElementType.Struct, _
                                                                                   LanguageElementType.Interface}
        Friend Function IsMemberOrType(ByVal CodeActive As LanguageElement) As Boolean
            Return MemberOrType.Contains(CodeActive.ElementType)
        End Function
        Friend Function HasExplicityVisibility(ByVal Element As AccessSpecifiedElement) As Boolean
            Return Element.VisibilityRange <> SourceRange.Empty
        End Function
#End Region

#Region "SA1100"
        Public Const Message_SA1100 As String = "SA1100 - Do not prefix calls with Base or MyBase unless a local implementation exists."
        Public Sub Available_SA1100(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1100(ea.CodeActive)
        End Sub
        Public Function Qualifies_SA1100(ByVal Element As IElement) As Boolean
            Dim Expression = TryCast(Element, BaseReferenceExpression)
            If Expression Is Nothing Then
                Return False
            End If
            Dim MethodReference = TryCast(Expression.Parent, MethodReferenceExpression)
            Dim BaseMethod = TryCast(MethodReference.GetDeclaration, Method)
            If BaseMethod Is Nothing Then
                Return False ' BaseMethod is non source code 
            End If
            Dim LocalMethods = MethodReference.GetClass.AllMethods().Cast(Of Method)()
            Dim LocalMethod = LocalMethods.Where(Function(m) m.IsOverride _
                                                 AndAlso XOverridesY(m, BaseMethod)).FirstOrDefault
            Return LocalMethod Is Nothing
        End Function
        Private Function XOverridesY(ByVal MethodX As Method, ByVal MethodY As Method) As Boolean
            ' X IsOverride
            If Not MethodX.IsOverride Then
                Return False
            End If
            ' Name of X = Name of Y
            If Not MethodX.Name = MethodY.Name Then
                Return False
            End If
            ' Class Prep
            Dim ClassX As [Class] = MethodX.GetClass
            Dim ClassY As [Class] = MethodY.GetClass
            ' ClassOf X descends from Class of Y
            If Not ClassX.DescendsFrom(ClassY) Then
                Return False
            End If
            ' Return Value of X = Return Value of Y
            If Not MethodX.MemberType = MethodY.MemberType Then
                Return False
            End If
            ' X Param Types = Y Param Types
            If Not ParamStringOf(MethodX) = ParamStringOf(MethodY) Then
                Return False
            End If
            Return True
        End Function
        Private Function ParamStringOf(ByVal Method As Method) As String
            Dim Result As String = String.Empty
            For Each Param As Parameter In Method.Parameters
                Result &= Param.Type.Name & "|"
            Next
            Return Result
        End Function

        Public Sub Fix_SA1100(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            ea.TextDocument.DeleteText(GetEndAdjustedRange(ea.CodeActive.Range, 0, 1))
        End Sub
        Private Function GetEndAdjustedRange(ByVal Range As SourceRange, ByVal Lines As Integer, ByVal Columns As Integer) As SourceRange
            Return New SourceRange(Range.Start, Range.End.OffsetPoint(Lines, Columns))
        End Function
#End Region
#Region "SA1101"
        Public Const Message_SA1101 As String = "SA1101 - Prefix local calls this or Me."
        Public Sub Available_SA1101(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1101(ea.CodeActive)
        End Sub
        Public Function Qualifies_SA1101(ByVal Element As IElement) As Boolean
            Dim MethodCall = TryCast(Element, MethodCall)
            If MethodCall Is Nothing Then
                Return False
            End If
            Dim Method = MethodCall.GetMethod()
            ' Is Methodcall Local
            If Not Method.GetClass Is MethodCall.GetClass Then
                Return False
            End If
            ' Is MethodCall Prefixed with ThisReference
            If Not MethodCall.Qualifier.IsIdenticalTo(New ThisReferenceExpression) Then
                Return False
            End If
            Return True
        End Function
        Public Sub Fix_SA1101(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            Dim NewCode As String = GenerateCode(New ThisReferenceExpression) + "."
            ea.TextDocument.InsertText(ea.CodeActive.Range.Start, NewCode)
        End Sub
#End Region
        '        Private Function GetRegionsInRange(ByVal sourceFile As SourceFile, ByVal parent As LanguageElement, ByVal memberRange As SourceRange, ByRef newRegionsToAdd As StringCollection) As RegionDirectiveCollection
        '            Dim parentRange As SourceRange = parent.Range
        '            Dim result As RegionDirectiveCollection = New RegionDirectiveCollection()
        '            For Each regionDirective As RegionDirective In CollectRegions(sourceFile.Regions)
        '                Dim regionDirectiveRange As SourceRange = regionDirective.Range
        '                If (Not parentRange.Contains(regionDirectiveRange)) Then
        '                    Continue For
        '                End If
        '                ' Region is inside the parent...
        '                Dim existingIndex As Integer = newRegionsToAdd.IndexOf(regionDirective.Name)
        '                If (existingIndex >= 0) Then ' Region already exists; no need to have it in the list of regions to add...
        '                    newRegionsToAdd.RemoveAt(existingIndex)
        '                End If
        '                If (memberRange.Contains(regionDirectiveRange)) Then
        '                    Continue For
        '                End If
        '                ' Region is a valid target.
        '                result.Add(regionDirective)
        '            Next
        '            Return result
        '        End Function
        '        Private Function CollectRegions(ByVal regions As RegionDirectiveCollection) As List(Of RegionDirective)
        '            Dim result = New List(Of RegionDirective)
        '            If (regions Is Nothing) Then
        '                Return result
        '            End If
        '            For Each regionDirective As RegionDirective In regions
        '                result.Add(regionDirective)
        '            Next
        '            Return result
        '        End Function
    End Module
End Namespace
