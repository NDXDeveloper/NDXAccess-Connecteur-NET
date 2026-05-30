Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>Tests d'intégration de la traduction des erreurs ACE en AccessQueryException.</summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class AccessErrorTranslationTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    Private Shared Function NewTable() As String
        Return "t_" & Guid.NewGuid().ToString("N")
    End Function

    <SkippableFact>
    Public Sub InvalidSql_ShouldThrowAccessQueryException()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            Dim act As Action = Sub() connection.ExecuteNonQuery("CECI N'EST PAS DU SQL VALIDE")
            act.Should().Throw(Of AccessQueryException)()
        End Using
    End Sub

    <SkippableFact>
    Public Sub DuplicatePrimaryKey_ShouldThrowTranslatedException()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim table = NewTable()

        Using connection = _fixture.CreateConnection()
            Try
                connection.ExecuteNonQuery($"CREATE TABLE {table} (code TEXT(10) CONSTRAINT pk PRIMARY KEY, libelle TEXT(50))")
                connection.ExecuteNonQuery($"INSERT INTO {table} (code, libelle) VALUES (?, ?)", "A", "Premier")

                Dim act As Action = Sub() connection.ExecuteNonQuery($"INSERT INTO {table} (code, libelle) VALUES (?, ?)", "A", "Doublon")

                ' Le provider ACE renvoie souvent un code natif 0 : on valide la traduction
                ' (AccessQueryException) et la préservation de l'exception OLE DB d'origine.
                Dim ex = act.Should().Throw(Of AccessQueryException)().Which
                ex.InnerException.Should().BeOfType(Of System.Data.OleDb.OleDbException)()
                ex.Message.Should().NotBeNullOrEmpty()
            Finally
                connection.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Sub

    <SkippableFact>
    Public Sub TranslateErrorsDisabled_ShouldThrowRawOleDbException()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim options = _fixture.GetOptions()
        options.TranslateErrors = False

        Using connection As IAccessConnection = New AccessConnection(options)
            Dim act As Action = Sub() connection.ExecuteNonQuery("SELECT * FROM table_qui_nexiste_pas_du_tout")
            act.Should().Throw(Of System.Data.OleDb.OleDbException)()
        End Using
    End Sub

End Class
