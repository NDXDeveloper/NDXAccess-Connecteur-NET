Imports System.Data.OleDb
Imports System.IO
Imports FluentAssertions
Imports Microsoft.Extensions.Logging
Imports NDXAccess
Imports NDXAccess.Examples
Imports Xunit

''' <summary>
''' Tests d'intégration des exemples « patterns réels » : DataAdapter (Fill/Update),
''' démarrage tout-en-un, pagination, DDL/relations, lecture CSV via ACE, logging.
''' </summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class MoreExamplesIntegrationTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    Private Shared Function NewTable() As String
        Return "t_" & Guid.NewGuid().ToString("N")
    End Function

    Private Shared Function TempDb() As String
        Return Path.Combine(Path.GetTempPath(), $"ndxaccess_ex_{Guid.NewGuid():N}.accdb")
    End Function

    Private Shared Sub CleanupFile(dbPath As String)
        Try
            OleDbConnection.ReleaseObjectPool()
            If File.Exists(dbPath) Then File.Delete(dbPath)
        Catch
        End Try
    End Sub

    <SkippableFact>
    Public Sub DataAdapter_FillModifyUpdate_ShouldPersistBatchChanges()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim table = NewTable()
        Using c = _fixture.CreateConnection()
            Try
                c.ExecuteNonQuery($"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, nom TEXT(100), email TEXT(200), actif YESNO)")
                c.ExecuteNonQuery($"INSERT INTO {table} (nom, email, actif) VALUES (?, ?, ?)", "Original", "o@example.com", True)

                Dim affected = DataAdapterExamples.FillModifyUpdate(c, table)
                affected.Should().BeGreaterThanOrEqualTo(2)   ' 1 update + 1 insert

                c.ExecuteScalar(Of Integer)($"SELECT COUNT(*) FROM {table}").Should().Be(2)
                c.ExecuteScalar(Of Integer)($"SELECT COUNT(*) FROM {table} WHERE nom = ?", "Nom modifié").Should().Be(1)
            Finally
                c.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Sub

    <SkippableFact>
    Public Async Function GettingStarted_RunAsync_ShouldCompleteFullFlow() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim dbPath = TempDb()
        Try
            Dim count = Await GettingStartedExample.RunAsync(dbPath)
            count.Should().BeGreaterThanOrEqualTo(1)
            File.Exists(dbPath).Should().BeTrue()
        Finally
            CleanupFile(dbPath)
        End Try
    End Function

    <SkippableFact>
    Public Sub Pagination_ShouldReturnDistinctPages()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim table = NewTable()
        Using c = _fixture.CreateConnection()
            Try
                c.ExecuteNonQuery($"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, v LONG)")
                For i = 1 To 10
                    c.ExecuteNonQuery($"INSERT INTO {table} (v) VALUES (?)", i)
                Next

                Dim page1 = PaginationExample.GetPage(c, table, "id", 3, 1)
                Dim page2 = PaginationExample.GetPage(c, table, "id", 3, 2)

                page1.Rows.Count.Should().Be(3)
                page2.Rows.Count.Should().Be(3)
                ' Les pages sont disjointes et ordonnées.
                Convert.ToInt32(page1.Rows(2)("id")).Should().BeLessThan(Convert.ToInt32(page2.Rows(0)("id")))
            Finally
                c.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Sub

    <SkippableFact>
    Public Sub SchemaDdl_ShouldEnforceUniqueAndForeignKey()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim dbPath = TempDb()
        Try
            AccessConnection.CreateDatabase(dbPath)
            Using c As IAccessConnection = New AccessConnection(New AccessConnectionOptions With {.DatabasePath = dbPath})
                SchemaDdlExample.CreateRelationalSchema(c)

                c.TableExists("clients").Should().BeTrue()
                c.TableExists("commandes").Should().BeTrue()

                c.ExecuteNonQuery("INSERT INTO clients (nom, email) VALUES (?, ?)", "Alice", "alice@example.com")
                Dim clientId = c.ExecuteScalar(Of Integer)("SELECT @@IDENTITY")

                ' Index unique sur email : doublon refusé.
                Dim dupAct As Action = Sub() c.ExecuteNonQuery("INSERT INTO clients (nom, email) VALUES (?, ?)", "Bob", "alice@example.com")
                dupAct.Should().Throw(Of AccessQueryException)()

                ' Clé étrangère : commande orpheline refusée.
                Dim orphanAct As Action = Sub() c.ExecuteNonQuery("INSERT INTO commandes (fkClient, montant) VALUES (?, ?)", 99999, 10D)
                orphanAct.Should().Throw(Of AccessQueryException)()

                ' Commande valide acceptée.
                c.ExecuteNonQuery("INSERT INTO commandes (fkClient, montant) VALUES (?, ?)", clientId, 10D).Should().Be(1)
            End Using
        Finally
            CleanupFile(dbPath)
        End Try
    End Sub

    <SkippableFact>
    Public Sub ExcelCsv_ReadCsv_ShouldReturnRows()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim folder = Path.Combine(Path.GetTempPath(), $"ndxcsv_{Guid.NewGuid():N}")
        Directory.CreateDirectory(folder)
        Dim csvName = "data.csv"
        File.WriteAllText(Path.Combine(folder, csvName),
            "nom,age" & Environment.NewLine & "Alice,30" & Environment.NewLine & "Bob,25" & Environment.NewLine)

        Try
            Dim dt = ExcelCsvExample.ReadCsv(folder, csvName)
            dt.Rows.Count.Should().Be(2)
            CStr(dt.Rows(0)("nom")).Should().Be("Alice")
        Finally
            Try
                OleDbConnection.ReleaseObjectPool()
                Directory.Delete(folder, True)
            Catch
            End Try
        End Try
    End Sub

    <SkippableFact>
    Public Sub Logging_ShouldCaptureConnectionActivity()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim logger As New CapturingLogger()
        Using conn = LoggingExample.CreateConnectionWithLogging(_fixture.DatabasePath, logger)
            conn.Open()
            conn.ExecuteScalar(Of Integer)("SELECT COUNT(*) FROM clients")
        End Using
        logger.Entries.Should().NotBeEmpty()
    End Sub

End Class

''' <summary>Logger de test capturant les messages émis par le connecteur.</summary>
Friend NotInheritable Class CapturingLogger
    Implements ILogger(Of AccessConnection)

    Public ReadOnly Entries As New List(Of String)()

    Public Function BeginScope(Of TState)(state As TState) As IDisposable Implements ILogger.BeginScope
        Return Nothing
    End Function

    Public Function IsEnabled(logLevel As LogLevel) As Boolean Implements ILogger.IsEnabled
        Return True
    End Function

    Public Sub Log(Of TState)(logLevel As LogLevel, eventId As EventId, state As TState, exception As Exception, formatter As Func(Of TState, Exception, String)) Implements ILogger.Log
        Entries.Add(formatter(state, exception))
    End Sub

End Class
