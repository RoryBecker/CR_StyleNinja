Imports System.ComponentModel
Namespace SA13XX
    Module Registration
        Public Sub RegisterRulesAndFixes(ByVal c As IContainer)
            RegisterNamingRules(c)
            RegisterRefactorings(c)
        End Sub
        Friend Sub RegisterNamingRules(ByVal c As IContainer)
            c.CreateIssue(Message_SA1302, AddressOf Qualifies_SA1302, SourceTypeEnum.Interface)
            'Call CreateIssue("SA1303", AddressOf SA1303_ConstantFieldsStartWithUpperCase_CheckCodeIssues)

            ' Rules: Require Uppercase 
            c.CreateIssue(Message_SA1300, AddressOf Qualifies_SA1300, SourceTypeEnum.MainElement)
            c.CreateIssue(Message_SA1304, AddressOf Qualifies_SA1304, SourceTypeEnum.Field)
            c.CreateIssue(Message_SA1307, AddressOf Qualifies_SA1307, SourceTypeEnum.Field)

            ' Rules: Require Lowercase
            c.CreateIssue(Message_SA1306, AddressOf Qualifies_SA1306, SourceTypeEnum.Field)

            c.CreateIssue(Message_SA1305, AddressOf Qualifies_SA1305, SourceTypeEnum.Variable)


            'Rules: Prefixes and Underscores.
            c.CreateIssue(Message_SA1308, AddressOf Qualifies_SA1308, SourceTypeEnum.Field)
            c.CreateIssue(Message_SA1309, AddressOf Qualifies_SA1309, SourceTypeEnum.Field)
            c.CreateIssue(Message_SA1310, AddressOf Qualifies_SA1310, SourceTypeEnum.Field)


            c.CreateIssue(Message_LocalsShouldStart, AddressOf Qualifies_LocalWithPoorPrefix, SourceTypeEnum.Local)
            c.CreateIssue(Message_FieldsShouldStart, AddressOf Qualifies_FieldWithPoorPrefix, SourceTypeEnum.Field)
            c.CreateIssue(Message_ParamsShouldStart, AddressOf Qualifies_ParamWithPoorPrefix, SourceTypeEnum.Param)
        End Sub

    End Module
End Namespace
