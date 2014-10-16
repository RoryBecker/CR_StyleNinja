Imports DevExpress.CodeRush.Core

Friend Module ExitStrategies
    Friend Function LocalPrefixNotSet() As Boolean
        Return CodeRush.CodeStyle.PrefixLocal = String.Empty
    End Function
    Friend Function FieldPrefixNotSet() As Boolean
        Return CodeRush.CodeStyle.PrefixField = String.Empty
    End Function
    Friend Function ParamPrefixNotSet() As Boolean
        Return CodeRush.CodeStyle.PrefixParam = String.Empty
    End Function

End Module
