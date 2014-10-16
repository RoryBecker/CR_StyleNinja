Imports System.ComponentModel
Imports DevExpress.CodeRush.StructuralParser
Imports DevExpress.CodeRush.Core
Imports System.Diagnostics.CodeAnalysis
Imports System.Drawing
Imports DevExpress.Refactor
Namespace SA15XX
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

#Region "SA1507"
        Public Const Message_SA1507 As String = "SA1507 - Multiple Blank Lines are Bad."
        Public Sub Available_SA1507(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1507(ea.Caret.Line)
        End Sub
        Public Function Qualifies_SA1507(ByVal Line As Integer) As Boolean
            Return Line > 1 AndAlso LineIsBlank(Line) AndAlso LineIsBlank(Line - 1)
        End Function
        Private Function LineIsBlank(ByVal Line As Integer) As Boolean
            Return CodeRush.Documents.GetLineAt(Line).Trim = String.Empty
        End Function
        Public Sub Fix_SA1507(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            ea.TextDocument.DeleteLine(ea.Caret.Line)
        End Sub
#End Region
    End Module

End Namespace
