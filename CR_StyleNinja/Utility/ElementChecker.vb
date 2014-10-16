Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports DevExpress.CodeRush.Core
Imports DevExpress.CodeRush.PlugInCore
Imports DevExpress.CodeRush.StructuralParser

Public Enum SourceTypeEnum
    Unknown
    [Class]
    Field
    Local
    LocalOrField
    Method
    Param
    Variable
    Member
    MainElement
    [Interface]
    VisibleItems
    Attribute
    MethodCall
    [Try]
    NamespaceReference
End Enum

Public Class ElementChecker
    Private mIssueMessage As String
    Private mQualifier As ElementQualifiesDelegate
    Private mSourceType As SourceTypeEnum
    Private mExitStrategy As Func(Of Boolean)
    Private mType As Type

    Public Delegate Function ElementQualifiesDelegate(ByVal Element As IElement) As Boolean
    Public Sub New(ByVal Qualifier As ElementQualifiesDelegate, ByVal IssueMessage As String, ByVal Type As Type, Optional ByVal ExitStrategy As Func(Of Boolean) = Nothing)
        mType = Type
        mSourceType = SourceTypeEnum.Unknown
        mIssueMessage = IssueMessage
        mQualifier = Qualifier
        mExitStrategy = ExitStrategy
    End Sub
    Public Sub New(ByVal Qualifier As ElementQualifiesDelegate, ByVal IssueMessage As String, ByVal SourceType As SourceTypeEnum, Optional ByVal ExitStrategy As Func(Of Boolean) = Nothing)
        mExitStrategy = ExitStrategy
        mSourceType = SourceType
        mIssueMessage = IssueMessage
        mQualifier = Qualifier
    End Sub

    Public Sub CheckCodeIssues(ByVal sender As Object, ByVal ea As CheckCodeIssuesEventArgs)
        If mExitStrategy IsNot Nothing AndAlso mExitStrategy.Invoke() Then
            Exit Sub
        End If
        If ea.Scope.ToLE IsNot Nothing Then
            Dim Finder = From e In GetFinder(ea.Scope) Where mQualifier.Invoke(e)
            For Each FoundItem In Finder
                'Todo: Find a way to provide a configurable custom color for StyleCop Violations
                ea.AddIssue(CodeIssueType.Hint, DecorateRange(FoundItem), mIssueMessage)
            Next
        End If
    End Sub
    Public Function DecorateRange(ByVal FoundItem As LanguageElement) As SourceRange
        Select Case FoundItem.ElementType
            Case LanguageElementType.Try
                Return New SourceRange(FoundItem.Range.Start, FoundItem.Range.Start.OffsetPoint(0, 3))
            Case LanguageElementType.NamespaceReference
                Return FoundItem.Range
            Case Else
                Return FoundItem.NameRange
        End Select
    End Function
    Private Function GetFinder(ByVal Scope As IElement, ByVal ElementType As IElement) As IEnumerable(Of LanguageElement)
        Return Elements(Scope.ToLE, ElementType.GetType)
    End Function
    Private Function GetFinder(ByVal Scope As IElement) As IEnumerable(Of LanguageElement)
        Select Case mSourceType
            Case SourceTypeEnum.Field
                Return Fields(Scope.ToLE)
            Case SourceTypeEnum.Local
                Return Locals(Scope.ToLE)
            Case SourceTypeEnum.LocalOrField
                Return FieldsOrLocals(Scope.ToLE)
            Case SourceTypeEnum.Variable
                Return Variables(Scope.ToLE)
            Case SourceTypeEnum.Param
                Return Params(Scope.ToLE)
            Case SourceTypeEnum.Member
                Return Members(Scope.ToLE)
            Case SourceTypeEnum.Attribute
                Return Attributes(Scope.ToLE)
            Case SourceTypeEnum.VisibleItems
                Return AccessSpecifiedItems(Scope.ToLE)
            Case SourceTypeEnum.Method
                Return Methods(Scope.ToLE)
            Case SourceTypeEnum.MainElement
                Return MainElements(Scope.ToLE)
            Case SourceTypeEnum.Interface
                Return Interfaces(Scope.ToLE)
            Case SourceTypeEnum.MethodCall
                Return MethodCall(Scope.ToLE)
            Case SourceTypeEnum.Try
                Return [Try](Scope.ToLE)
            Case SourceTypeEnum.Class
                Return Classes(Scope.ToLE)
            Case SourceTypeEnum.NamespaceReference
                Return Elements(Scope.ToLE)
            Case Else
                Return Elements(Scope.ToLE)
        End Select
        'Dim X = Me.GetFinder(Scope)
        'Dim Y = GetFinder(Scope)

    End Function

End Class