Imports System.Data.OleDb
Imports System.IO
Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>
''' Tests d'intégration du compactage (CompactDatabase / CompactDatabaseAsync).
''' Nécessite le moteur DAO installé avec ACE. Ignorés si ACE absent.
''' </summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class AccessMaintenanceTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    Private Shared Function NewTable() As String
        Return "t_" & Guid.NewGuid().ToString("N")
    End Function

    <SkippableFact>
    Public Async Function CompactDatabaseAsync_ShouldPreserveData() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim table = NewTable()

        Using connection = _fixture.CreateConnection()
            Try
                Await connection.ExecuteNonQueryAsync($"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, valeur LONG)")
                For i = 1 To 50
                    Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (valeur) VALUES (?)", {i})
                Next

                ' Compactage en place : ferme la connexion, compacte via DAO, puis on rouvre.
                Await connection.CompactDatabaseAsync()

                File.Exists(_fixture.DatabasePath).Should().BeTrue()

                Dim count = Await connection.ExecuteScalarAsync(Of Integer)($"SELECT COUNT(*) FROM {table}")
                count.Should().Be(50)
            Finally
                connection.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Function

    <SkippableFact>
    Public Sub CompactDatabase_DuringTransaction_ShouldThrow()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            connection.BeginTransaction()
            Try
                Dim act As Action = Sub() connection.CompactDatabase()
                act.Should().Throw(Of InvalidOperationException)()
            Finally
                connection.Rollback()
            End Try
        End Using
    End Sub

    <SkippableFact>
    Public Sub CompactDatabase_ToNewFile_ShouldCreateCompactedCopy()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        ' Scénario réaliste : compacter une base existante vers un fichier de sauvegarde.
        Dim target = Path.Combine(Path.GetTempPath(), $"ndxaccess_dst_{Guid.NewGuid():N}.accdb")

        Try
            Using connection = _fixture.CreateConnection()
                ' Compactage SYNCHRONE de la base de la fixture vers un NOUVEAU fichier.
                connection.CompactDatabase(target)
            End Using

            File.Exists(target).Should().BeTrue()
            File.Exists(_fixture.DatabasePath).Should().BeTrue()   ' l'original reste présent

            ' Le fichier compacté contient bien la table seedée 'clients' (3 lignes).
            Using c2 As IAccessConnection = New AccessConnection(New AccessConnectionOptions With {.DatabasePath = target})
                c2.ExecuteScalar(Of Integer)("SELECT COUNT(*) FROM clients").Should().Be(3)
            End Using
        Finally
            Try
                OleDbConnection.ReleaseObjectPool()
                If File.Exists(target) Then File.Delete(target)
            Catch
            End Try
        End Try
    End Sub

End Class
