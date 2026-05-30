Imports System.IO
Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>Tests d'intégration des helpers de schéma et de CreateDatabase.</summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class AccessSchemaTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    <SkippableFact>
    Public Sub TableExists_ShouldReturnTrueForSeededTable()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Using connection = _fixture.CreateConnection()
            connection.TableExists("clients").Should().BeTrue()
            connection.TableExists("CLIENTS").Should().BeTrue()           ' insensible à la casse
            connection.TableExists("table_inexistante").Should().BeFalse()
        End Using
    End Sub

    <SkippableFact>
    Public Async Function GetTableNamesAsync_ShouldContainSeededTable() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Using connection = _fixture.CreateConnection()
            Dim names = Await connection.GetTableNamesAsync()
            names.Should().Contain("clients")
        End Using
    End Function

    <SkippableFact>
    Public Sub GetColumns_ShouldReturnColumnSchema()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Using connection = _fixture.CreateConnection()
            Dim columns = connection.GetColumns("clients")
            columns.Rows.Count.Should().BeGreaterThan(0)
        End Using
    End Sub

    <SkippableFact>
    Public Sub GetQueryNames_ShouldListSavedQuery()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim query = "qschema_" & Guid.NewGuid().ToString("N")
        Using connection = _fixture.CreateConnection()
            Try
                connection.ExecuteNonQuery($"CREATE PROCEDURE {query} AS SELECT id FROM clients")
                connection.GetQueryNames().Should().Contain(query)
            Finally
                Try : connection.ExecuteNonQuery($"DROP PROCEDURE {query}") : Catch : End Try
            End Try
        End Using
    End Sub

    <SkippableFact>
    Public Async Function CreateDatabase_ShouldCreateUsableFile() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim dbPath = Path.Combine(Path.GetTempPath(), $"ndxaccess_new_{Guid.NewGuid():N}.accdb")

        Try
            Await AccessConnection.CreateDatabaseAsync(dbPath)
            File.Exists(dbPath).Should().BeTrue()

            Using connection As IAccessConnection = New AccessConnection(New AccessConnectionOptions With {.DatabasePath = dbPath})
                Await connection.ExecuteNonQueryAsync("CREATE TABLE t (id AUTOINCREMENT PRIMARY KEY, v LONG)")
                Await connection.ExecuteNonQueryAsync("INSERT INTO t (v) VALUES (?)", {42})
                Dim n = Await connection.ExecuteScalarAsync(Of Integer)("SELECT COUNT(*) FROM t")
                n.Should().Be(1)
            End Using
        Finally
            Try
                System.Data.OleDb.OleDbConnection.ReleaseObjectPool()
                If File.Exists(dbPath) Then File.Delete(dbPath)
            Catch
            End Try
        End Try
    End Function

    <SkippableFact>
    Public Sub CreateDatabase_WhenFileExists_ShouldThrow()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim act As Action = Sub() AccessConnection.CreateDatabase(_fixture.DatabasePath)
        act.Should().Throw(Of IOException)()
    End Sub

End Class
