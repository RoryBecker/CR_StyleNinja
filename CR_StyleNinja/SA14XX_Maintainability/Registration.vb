Imports System.ComponentModel
Imports DevExpress.CodeRush.StructuralParser
Imports DevExpress.CodeRush.Core
Namespace SA14XX
    Module Registration
        Public Sub RegisterRulesAndFixes(ByVal C As IContainer)
            RegisterRules(C)
            RegisterFixes(C)
        End Sub
        Private Sub RegisterRules(ByVal C As IContainer)
            C.CreateIssue(Message_SA1400, AddressOf Qualifies_SA1400, SourceTypeEnum.VisibleItems)
            C.CreateIssue(Message_SA1401, AddressOf Qualifies_SA1401, SourceTypeEnum.Field)
            C.CreateIssue(Message_SA1402, AddressOf Qualifies_SA1402, SourceTypeEnum.Class)
            C.CreateIssue(Message_SA1403, AddressOf Qualifies_SA1403, SourceTypeEnum.Attribute)
            C.CreateIssue(Message_SA1404, AddressOf Qualifies_SA1404, SourceTypeEnum.Attribute)
            C.CreateIssue(Message_SA1405, AddressOf Qualifies_SA1405, SourceTypeEnum.MethodCall)
            C.CreateIssue(Message_SA1409, AddressOf Qualifies_SA1409, SourceTypeEnum.Try)
        End Sub

        Private Sub RegisterFixes(ByVal C As IContainer)
            ' SA1400
            C.CreateRefactoring("MakeVisibilityExplicit", "Make Visibility Explicit", AddressOf SA1400_Available, AddressOf Fix_SA1400).SolvedIssues.Add(Message_SA1400)
            ' SA1401
            C.CreateCodeProvider("MakeFieldPrivate", "Make Field Private", AddressOf Available_SA1401, AddressOf Fix_SA1401).SolvedIssues.Add(Message_SA1401)
            ' SA1409
            C.CreateCodeProvider("RemoveTryX", "Remove Try..X", AddressOf Available_SA1409, AddressOf Fix_SA1409).SolvedIssues.Add(Message_SA1409)
        End Sub
    End Module
End Namespace
