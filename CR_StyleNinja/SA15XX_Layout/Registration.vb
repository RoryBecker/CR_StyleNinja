Imports System.ComponentModel
Imports DevExpress.CodeRush.StructuralParser
Imports DevExpress.CodeRush.Core
Imports System.Runtime.CompilerServices

Namespace SA15XX
    Module Registration
        Public Sub RegisterRulesAndFixes(ByVal C As IContainer)
            RegisterRules(C)
            RegisterFixes(C)
        End Sub
        Private Sub RegisterRules(ByVal C As IContainer)
            C.CreateIssue(Message_SA1507, AddressOf Qualifies_SA1507)
        End Sub
        Private Sub RegisterFixes(ByVal C As IContainer)
            ' SA1507
            C.CreateRefactoring("Remove Blank Line",
                                "Remove Blank Line",
                                AddressOf Available_SA1507,
                                AddressOf Fix_SA1507).SolvedIssues.Add(Message_SA1507)
        End Sub
        '' Please ensure the following line is not missing from your plugin's InitializeComponent
        '' components = New System.ComponentModel.Container()
        'Public Sub CreateTooManyBlankLines()
        '    Dim TooManyBlankLines As New DevExpress.CodeRush.Core.IssueProvider(components)
        '    CType(TooManyBlankLines, System.ComponentModel.ISupportInitialize).BeginInit()
        '    TooManyBlankLines.ProviderName = "IssueProviderName" ' Should be Unique
        '    TooManyBlankLines.DisplayName = "IssueDisplayName"
        '    AddHandler TooManyBlankLines.CheckCodeIssues, AddressOf TooManyBlankLines_CheckCodeIssues
        '    CType(TooManyBlankLines, System.ComponentModel.ISupportInitialize).EndInit()
        'End Sub
        'Private Sub TooManyBlankLines_CheckCodeIssues(ByVal sender As Object, ByVal ea As CheckCodeIssuesEventArgs)
        '    ' This method is executed when the system checks for your issue.
        '    Dim Scope = TryCast(ea.Scope, LanguageElement)
        '    If Scope Is Nothing Then
        '        Exit Sub
        '    End If
        '    Dim Finder As New ElementEnumerable(Scope, GetType(DevExpress.CodeRush.StructuralParser.Method), True)
        '    For Each FoundItem As DevExpress.CodeRush.StructuralParser.Method In Finder
        '        ' ea.AddError(FoundItem.NameRange, "Provide Error Detail here.")
        '        ' ea.AddHint(FoundItem.NameRange, "Provide hint here.")
        '        ' You get the general drift. :)

        '    Next
        'End Sub
    End Module
    Public Module SourcePointExt
        <Extension()> _
        Public Function ToRange(ByVal Source As SourcePoint, Optional ByVal Width As Integer = 0) As SourceRange
            Return New SourceRange(Source, Source.OffsetPoint(0, Width))
        End Function
    End Module
End Namespace
