Imports System.ComponentModel
Imports DevExpress.CodeRush.StructuralParser
Imports DevExpress.CodeRush.Core
Namespace SA12XX
    Module Registration
        Public Sub RegisterRulesAndFixes(ByVal C As IContainer)
            RegisterRules(C)
            RegisterFixes(C)
        End Sub
        Private Sub RegisterRules(ByVal C As IContainer)
            C.CreateIssue(Message_SA1200, AddressOf Qualifies_SA1200, SourceTypeEnum.NamespaceReference)
        End Sub
        Private Sub RegisterFixes(ByVal C As IContainer)
            ' SA1200
            Dim Refactoring = C.CreateRefactoring("Move Using/Imports inside Namespace",
                                                  "Move Using/Imports inside Namespace",
                                                  AddressOf Available_SA1200,
                                                  AddressOf Fix_SA1200)
            Refactoring.SolvedIssues.Add(Message_SA1200)


        End Sub
    End Module
End Namespace
