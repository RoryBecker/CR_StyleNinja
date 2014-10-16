Imports DevExpress.CodeRush.StructuralParser
Imports DevExpress.CodeRush.Core
Namespace SA13XX
    Module Renaming
        Public Sub RenameElement(ByVal Declaration As LanguageElement, ByVal NewName As String)
            ' Get References
            Dim DocumentDict = GetReferenceRanges(Declaration, True)
            'Rename References
            CodeRush.VSSettings.SuppressVBPrettyListing()
            For Each Entry In DocumentDict
                CodeRush.File.ChangeFile(Entry.Key, Entry.Value.ToArray, NewName)
            Next
            CodeRush.VSSettings.RestoreVBPrettyListing()
        End Sub
        Private Function GetReferenceRanges(ByVal Declaration As LanguageElement, ByVal IncludeDeclaration As Boolean) As Dictionary(Of SourceFile, List(Of SourceRange))
            Dim Documents As New Dictionary(Of SourceFile, List(Of SourceRange))
            Dim Found = CodeRush.Refactoring.FindAllReferences(Declaration.Solution, Declaration).ToLanguageElementCollection
            Dim FoundLE = Found.OfType(Of LanguageElement)()
            ' Add Declaration NameRange
            If IncludeDeclaration Then
                Documents(Declaration.GetSourceFile) = New List(Of SourceRange)
                Documents(Declaration.GetSourceFile).Add(Declaration.NameRange)
            End If
            For Each Reference In FoundLE
                Dim SourceFile = Reference.GetSourceFile
                If Not Documents.Keys.Contains(SourceFile) Then
                    Documents(SourceFile) = New List(Of SourceRange)
                End If
                Documents(SourceFile).Add(Reference.NameRange)
            Next
            Return Documents
        End Function
    End Module
End Namespace
