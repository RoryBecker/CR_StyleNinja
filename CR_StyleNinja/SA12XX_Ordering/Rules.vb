Imports System.ComponentModel
Imports DevExpress.CodeRush.StructuralParser
Imports DevExpress.CodeRush.Core
Imports System.Diagnostics.CodeAnalysis
Imports System.Runtime.CompilerServices
Namespace SA12XX
    Friend Module Rules
#Region "SA1200"
        Public Const Message_SA1200 As String = "SA1200: Do not place using or imports directives outside of a namespace element"
        Public Sub Available_SA1200(ByVal sender As Object, ByVal ea As CheckContentAvailabilityEventArgs)
            ea.Available = Qualifies_SA1200(ea.CodeActive)
        End Sub
        Public Function Qualifies_SA1200(ByVal Element As IElement) As Boolean
            If CodeRush.Language.GetLanguageID(Element) <> "CSharp" Then
                Return False
            End If
            Element = Element.GetParent(LanguageElementType.NamespaceReference)
            If Element Is Nothing Then
                Return False
            End If
            Dim IsReference = Element.ElementType = LanguageElementType.NamespaceReference
            Dim PoorlyParented = Element.Parent IsNot Nothing AndAlso Element.Parent.ElementType <> LanguageElementType.Namespace

            Return IsReference AndAlso PoorlyParented
        End Function

        Public Sub Fix_SA1200(ByVal sender As Object, ByVal ea As ApplyContentEventArgs)
            Dim ActiveDoc As TextDocument = CodeRush.Documents.ActiveTextDocument
            Dim SourceFile As SourceFile = ea.Element.GetSourceFile
            Dim NamespaceReferences = Elements(SourceFile, GetType(NamespaceReference), False)
            'Dim ReversedReferences = From Item In NameSpaceReferences Select Item.GenerateCode
            Dim NamespaceBlockRange As SourceRange = GetCombinedBlockRange(NamespaceReferences)
            For Each TheNamespace As [Namespace] In Elements(SourceFile, GetType([Namespace]))
                Dim PointInsideNameSpace = TheNamespace.BlockCodeRange.Start
                'Dim NamespaceReferenceCode As String = NamespaceReference.GenerateCode
                Dim NamespaceReferenceCode As String = ActiveDoc.GetText(NamespaceBlockRange)
                ActiveDoc.QueueInsert(PointInsideNameSpace, NamespaceReferenceCode)
            Next
            ActiveDoc.QueueDelete(NamespaceBlockRange)
            ActiveDoc.ApplyQueuedEdits("Moved Namespace References", True)
        End Sub
        Private Function GetCombinedBlockRange(ByVal NamespaceReferences As System.Collections.Generic.IEnumerable(Of LanguageElement)) As SourceRange
            Dim Result As SourceRange
            Result.Start = NamespaceReferences.First.Range.Start
            Result.End = NamespaceReferences.Last.Range.Start.OffsetPoint(1, 0)
            Return Result
        End Function

#End Region
    End Module
End Namespace

