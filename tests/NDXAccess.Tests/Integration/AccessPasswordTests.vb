Imports System.Data.OleDb
Imports System.IO
Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>Tests d'intégration des bases protégées par mot de passe.</summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class AccessPasswordTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    <SkippableFact>
    Public Sub PasswordProtectedDatabase_OpenWithCorrectAndWrongPassword()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim dbPath = Path.Combine(Path.GetTempPath(), $"ndxaccess_pwd_{Guid.NewGuid():N}.accdb")
        Const pwd As String = "S3cret!ACE"

        Try
            ' Création d'une base protégée.
            AccessConnection.CreateDatabase(dbPath, pwd)

            ' Bon mot de passe -> tout fonctionne.
            Using ok As IAccessConnection = New AccessConnection(New AccessConnectionOptions With {.DatabasePath = dbPath, .Password = pwd})
                ok.ExecuteNonQuery("CREATE TABLE t (id AUTOINCREMENT PRIMARY KEY, v LONG)")
                ok.ExecuteNonQuery("INSERT INTO t (v) VALUES (?)", 1)
                ok.ExecuteScalar(Of Integer)("SELECT COUNT(*) FROM t").Should().Be(1)
            End Using
            OleDbConnection.ReleaseObjectPool()

            ' Mauvais mot de passe -> échec traduit.
            Using bad As IAccessConnection = New AccessConnection(New AccessConnectionOptions With {.DatabasePath = dbPath, .Password = "mauvais"})
                Dim act As Action = Sub() bad.Open()
                act.Should().Throw(Of AccessQueryException)()
            End Using
            OleDbConnection.ReleaseObjectPool()

            ' Sans mot de passe -> échec également.
            Using none As IAccessConnection = New AccessConnection(New AccessConnectionOptions With {.DatabasePath = dbPath})
                Dim act As Action = Sub() none.Open()
                act.Should().Throw(Of AccessQueryException)()
            End Using
        Finally
            Try
                OleDbConnection.ReleaseObjectPool()
                If File.Exists(dbPath) Then File.Delete(dbPath)
            Catch
            End Try
        End Try
    End Sub

End Class
