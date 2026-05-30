Imports System.Collections.Generic
Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>Tests d'intégration : BulkInsert, mapping Query(Of T), paramètres nommés.</summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class AccessBulkAndMappingTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    Private Shared Function NewTable() As String
        Return "t_" & Guid.NewGuid().ToString("N")
    End Function

    ''' <summary>Classe cible du mapping (propriétés = colonnes).</summary>
    Public Class Client
        Public Property Id As Integer
        Public Property Nom As String
        Public Property Email As String
        Public Property Actif As Boolean
    End Class

    <SkippableFact>
    Public Async Function BulkInsertAsync_ShouldInsertManyRows() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim table = NewTable()

        Using connection = _fixture.CreateConnection()
            Try
                Await connection.ExecuteNonQueryAsync($"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, nom TEXT(50), montant CURRENCY)")

                Dim rows As New List(Of Object())()
                For i = 1 To 200
                    rows.Add(New Object() {$"Nom{i}", CDec(i) * 1.5D})
                Next

                Dim inserted = Await connection.BulkInsertAsync(table, {"nom", "montant"}, rows)
                inserted.Should().Be(200)

                Dim count = Await connection.ExecuteScalarAsync(Of Integer)($"SELECT COUNT(*) FROM {table}")
                count.Should().Be(200)
            Finally
                connection.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Function

    <SkippableFact>
    Public Sub BulkInsert_Sync_ShouldInsertRows()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim table = NewTable()

        Using connection = _fixture.CreateConnection()
            Try
                connection.ExecuteNonQuery($"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, v LONG)")
                Dim rows As New List(Of Object())() From {New Object() {1}, New Object() {2}, New Object() {3}}
                connection.BulkInsert(table, {"v"}, rows).Should().Be(3)
            Finally
                connection.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Sub

    <SkippableFact>
    Public Async Function ExecuteQueryOfT_ShouldMapRowsToObjects() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            Dim clients = Await connection.ExecuteQueryAsync(Of Client)("SELECT id, nom, email, actif FROM clients ORDER BY nom")
            clients.Should().HaveCount(3)
            clients(0).Nom.Should().Be("Alice")
            clients(0).Email.Should().Be("alice@example.com")
            clients(0).Actif.Should().BeTrue()
            clients(1).Nom.Should().Be("Bob")
            clients(1).Actif.Should().BeFalse()
        End Using
    End Function

    <SkippableFact>
    Public Sub ExecuteQueryOfT_Sync_ShouldMapScalarColumns()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            Dim actifs = connection.ExecuteQuery(Of Client)("SELECT id, nom FROM clients WHERE actif = ? ORDER BY nom", True)
            actifs.Should().HaveCount(2)
            actifs(0).Nom.Should().Be("Alice")
            actifs(1).Nom.Should().Be("Charlie")
            actifs(0).Email.Should().BeNull()   ' colonne non sélectionnée -> défaut
        End Using
    End Sub

    <SkippableFact>
    Public Async Function ExecuteQueryNamedAsync_ShouldUseNamedParameters() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            Dim params As New Dictionary(Of String, Object) From {{"actif", True}}
            Dim dt = Await connection.ExecuteQueryNamedAsync("SELECT nom FROM clients WHERE actif = @actif ORDER BY nom", params)
            dt.Rows.Count.Should().Be(2)
            CStr(dt.Rows(0)("nom")).Should().Be("Alice")
        End Using
    End Function

    <SkippableFact>
    Public Sub ExecuteScalarNamed_ShouldUseNamedParameters()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            Dim params As New Dictionary(Of String, Object) From {{"actif", True}}
            Dim count = connection.ExecuteScalarNamed(Of Integer)("SELECT COUNT(*) FROM clients WHERE actif = @actif", params)
            count.Should().Be(2)
        End Using
    End Sub

End Class
