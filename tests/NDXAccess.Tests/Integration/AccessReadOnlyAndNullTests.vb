Imports System.Data.OleDb
Imports System.IO
Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>Tests d'intégration : mode lecture seule et gestion des valeurs NULL.</summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class AccessReadOnlyAndNullTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    Public Class NullRow
        Public Property Nom As String
        Public Property Valeur As Integer
    End Class

    Private Shared Function TempDb() As String
        Return Path.Combine(Path.GetTempPath(), $"ndxaccess_t_{Guid.NewGuid():N}.accdb")
    End Function

    Private Shared Sub Cleanup(dbPath As String)
        Try
            OleDbConnection.ReleaseObjectPool()
            If File.Exists(dbPath) Then File.Delete(dbPath)
        Catch
        End Try
    End Sub

    <SkippableFact>
    Public Sub ReadOnlyMode_ShouldAllowSelectButRejectWrite()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim dbPath = TempDb()

        Try
            AccessConnection.CreateDatabase(dbPath)
            Using w As IAccessConnection = New AccessConnection(New AccessConnectionOptions With {.DatabasePath = dbPath})
                w.ExecuteNonQuery("CREATE TABLE t (id AUTOINCREMENT PRIMARY KEY, v LONG)")
                w.ExecuteNonQuery("INSERT INTO t (v) VALUES (?)", 10)
            End Using
            OleDbConnection.ReleaseObjectPool()

            ' Lecture seule : SELECT OK, écriture refusée.
            Using r As IAccessConnection = New AccessConnection(New AccessConnectionOptions With {.DatabasePath = dbPath, .OpenReadOnly = True})
                r.ExecuteScalar(Of Integer)("SELECT COUNT(*) FROM t").Should().Be(1)
                Dim act As Action = Sub() r.ExecuteNonQuery("INSERT INTO t (v) VALUES (?)", 20)
                act.Should().Throw(Of AccessQueryException)()
            End Using
            OleDbConnection.ReleaseObjectPool()

            ' DatabaseInfo.IsReadOnly reflète le mode.
            Dim factory = New AccessConnectionFactory(New AccessConnectionOptions With {.DatabasePath = dbPath, .OpenReadOnly = True})
            Dim info = New AccessHealthCheck(factory).GetDatabaseInfo()
            info.IsReadOnly.Should().BeTrue()
        Finally
            Cleanup(dbPath)
        End Try
    End Sub

    <SkippableFact>
    Public Sub NullValues_ShouldMapToDefaults()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim dbPath = TempDb()

        Try
            AccessConnection.CreateDatabase(dbPath)
            Using c As IAccessConnection = New AccessConnection(New AccessConnectionOptions With {.DatabasePath = dbPath})
                c.ExecuteNonQuery("CREATE TABLE t (id AUTOINCREMENT PRIMARY KEY, nom TEXT(50), valeur LONG)")
                c.ExecuteNonQuery("INSERT INTO t (nom) VALUES (?)", "seul")   ' valeur = NULL
                c.ExecuteNonQuery("INSERT INTO t (valeur) VALUES (?)", 5)     ' nom = NULL

                ' Scalaire sur une valeur NULL -> défaut du type.
                c.ExecuteScalar(Of Integer)("SELECT valeur FROM t WHERE nom = ?", "seul").Should().Be(0)
                c.ExecuteScalar(Of String)("SELECT nom FROM t WHERE valeur = ?", 5).Should().BeNull()

                ' Mapping objet : NULL -> propriété laissée à sa valeur par défaut.
                Dim rows = c.ExecuteQuery(Of NullRow)("SELECT nom, valeur FROM t ORDER BY id")
                rows.Should().HaveCount(2)
                rows(0).Nom.Should().Be("seul")
                rows(0).Valeur.Should().Be(0)        ' NULL -> 0
                rows(1).Nom.Should().BeNull()        ' NULL -> Nothing
                rows(1).Valeur.Should().Be(5)
            End Using
        Finally
            Cleanup(dbPath)
        End Try
    End Sub

End Class
