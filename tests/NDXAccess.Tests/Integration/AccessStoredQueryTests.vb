Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>
''' Tests d'intégration des requêtes enregistrées (stored queries) Access.
''' Rappel : Access n'accepte que des paramètres d'ENTRÉE (pas de OUT/INOUT).
''' </summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class AccessStoredQueryTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    Private Shared Function Unique(prefix As String) As String
        Return prefix & Guid.NewGuid().ToString("N")
    End Function

    <SkippableFact>
    Public Async Function StoredQuery_SelectWithInputParameter_ShouldReturnFilteredRows() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim table = Unique("t_")
        Dim query = Unique("q_")

        Using connection = _fixture.CreateConnection()
            Try
                Await connection.ExecuteNonQueryAsync(
                    $"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, nom TEXT(50), categorie TEXT(20))")
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (nom, categorie) VALUES (?, ?)", {"Alpha", "A"})
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (nom, categorie) VALUES (?, ?)", {"Beta", "B"})
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (nom, categorie) VALUES (?, ?)", {"Gamma", "A"})

                ' Création d'une requête enregistrée paramétrée (paramètre IN uniquement).
                Await connection.ExecuteNonQueryAsync(
                    $"CREATE PROCEDURE {query} (prmCat TEXT) AS SELECT nom FROM {table} WHERE categorie = prmCat ORDER BY nom")

                Dim result = Await connection.ExecuteStoredQueryAsync(query, {"A"})

                result.Rows.Count.Should().Be(2)
                CStr(result.Rows(0)("nom")).Should().Be("Alpha")
                CStr(result.Rows(1)("nom")).Should().Be("Gamma")
            Finally
                Try : connection.ExecuteNonQuery($"DROP PROCEDURE {query}") : Catch : End Try
                connection.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Function

    <SkippableFact>
    Public Async Function StoredQuery_ActionQuery_ShouldReturnAffectedRows() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim table = Unique("t_")
        Dim query = Unique("q_")

        Using connection = _fixture.CreateConnection()
            Try
                Await connection.ExecuteNonQueryAsync(
                    $"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, nom TEXT(50), categorie TEXT(20))")
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (nom, categorie) VALUES (?, ?)", {"Alpha", "A"})
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (nom, categorie) VALUES (?, ?)", {"Beta", "B"})
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (nom, categorie) VALUES (?, ?)", {"Gamma", "A"})

                Await connection.ExecuteNonQueryAsync(
                    $"CREATE PROCEDURE {query} (prmCat TEXT) AS DELETE FROM {table} WHERE categorie = prmCat")

                Dim affected = Await connection.ExecuteStoredQueryNonQueryAsync(query, {"A"})
                affected.Should().Be(2)

                Dim remaining = Await connection.ExecuteScalarAsync(Of Integer)($"SELECT COUNT(*) FROM {table}")
                remaining.Should().Be(1)
            Finally
                Try : connection.ExecuteNonQuery($"DROP PROCEDURE {query}") : Catch : End Try
                connection.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Function

End Class
