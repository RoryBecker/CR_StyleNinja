Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports DevExpress.CodeRush.Core
Imports DevExpress.CodeRush.PlugInCore
Imports DevExpress.CodeRush.StructuralParser
Imports System.Runtime.CompilerServices

Public Module LanguageElementExt
    Private ClassOrStruct As LanguageElementType() = New LanguageElementType() {LanguageElementType.Class, LanguageElementType.Struct}
    <Extension()> _
    Public Function ToLE(ByVal Source As IElement) As LanguageElement
        Return TryCast(Source, LanguageElement)
    End Function

    <Extension()> _
    Friend Function IsHungarian(ByVal Source As String) As Boolean
        Return Source.Length > 2 andalso (Char.IsLower(Source.First) AndAlso Char.IsUpper(Source.Skip(1).First)) _
        OrElse Source.Length > 3 andalso (Char.IsLower(Source.First) AndAlso Char.IsLower(Source.Skip(1).First) AndAlso Char.IsUpper(Source.Skip(2).First))
    End Function
    <Extension()> _
    Public Function GenerateCode(ByVal Source As LanguageElement) As String
        Return CodeRush.CodeMod.GenerateCode(Source)
    End Function
End Module