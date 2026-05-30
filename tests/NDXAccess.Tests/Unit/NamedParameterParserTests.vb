Imports System.Collections.Generic
Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>Tests unitaires du parseur de paramètres nommés (@nom -> ?).</summary>
<Trait("Category", "Unit")>
Public Class NamedParameterParserTests

    <Fact>
    Public Sub Translate_SimpleNamedParams_ShouldProducePositional()
        Dim params As New Dictionary(Of String, Object) From {{"nom", "Alice"}, {"actif", True}}
        Dim result = NamedParameterParser.Translate("SELECT * FROM clients WHERE nom = @nom AND actif = @actif", params)

        result.Sql.Should().Be("SELECT * FROM clients WHERE nom = ? AND actif = ?")
        result.Values.Should().Equal(New Object() {"Alice", True})
    End Sub

    <Fact>
    Public Sub Translate_RepeatedParam_ShouldAddValueEachOccurrence()
        Dim params As New Dictionary(Of String, Object) From {{"x", 5}}
        Dim result = NamedParameterParser.Translate("SELECT * FROM t WHERE a = @x OR b = @x", params)

        result.Sql.Should().Be("SELECT * FROM t WHERE a = ? OR b = ?")
        result.Values.Should().Equal(New Object() {5, 5})
    End Sub

    <Fact>
    Public Sub Translate_ShouldIgnoreAtInsideStringLiterals()
        Dim params As New Dictionary(Of String, Object) From {{"v", "x"}}
        Dim result = NamedParameterParser.Translate("SELECT * FROM t WHERE email = '@notaparam' AND code = @v", params)

        result.Sql.Should().Be("SELECT * FROM t WHERE email = '@notaparam' AND code = ?")
        result.Values.Should().Equal(New Object() {"x"})
    End Sub

    <Fact>
    Public Sub Translate_ShouldLeaveUnknownTokensLikeIdentity()
        Dim params As New Dictionary(Of String, Object) From {{"id", 1}}
        Dim result = NamedParameterParser.Translate("SELECT @@IDENTITY FROM t WHERE id = @id", params)

        result.Sql.Should().Be("SELECT @@IDENTITY FROM t WHERE id = ?")
        result.Values.Should().Equal(New Object() {1})
    End Sub

    <Fact>
    Public Sub Translate_NamesWithAtPrefixInDictionary_ShouldAlsoMatch()
        Dim params As New Dictionary(Of String, Object) From {{"@id", 7}}
        Dim result = NamedParameterParser.Translate("SELECT * FROM t WHERE id = @id", params)

        result.Sql.Should().Be("SELECT * FROM t WHERE id = ?")
        result.Values.Should().Equal(New Object() {7})
    End Sub

    <Fact>
    Public Sub Translate_NoParams_ShouldReturnSqlUnchanged()
        Dim result = NamedParameterParser.Translate("SELECT 1 FROM t", Nothing)
        result.Sql.Should().Be("SELECT 1 FROM t")
        result.Values.Should().BeEmpty()
    End Sub

End Class
