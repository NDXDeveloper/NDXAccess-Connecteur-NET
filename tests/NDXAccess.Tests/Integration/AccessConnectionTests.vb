Imports System.Data
Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>
''' Tests d'intégration de AccessConnection sur une base .accdb réelle.
''' Ignorés automatiquement si le provider ACE n'est pas disponible.
''' Note : le nettoyage (DROP TABLE) est effectué en Finally de façon SYNCHRONE,
''' car VB.NET interdit Await dans un bloc Finally.
''' </summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class AccessConnectionTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    Private Shared Function NewTable() As String
        Return "t_" & Guid.NewGuid().ToString("N")
    End Function

    <SkippableFact>
    Public Sub Open_ShouldOpenConnection()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            connection.Open()
            connection.State.Should().Be(ConnectionState.Open)
        End Using
    End Sub

    <SkippableFact>
    Public Async Function OpenAsync_ThenCloseAsync_ShouldToggleState() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            Await connection.OpenAsync()
            connection.State.Should().Be(ConnectionState.Open)
            Await connection.CloseAsync()
            connection.State.Should().Be(ConnectionState.Closed)
        End Using
    End Function

    <SkippableFact>
    Public Async Function ExecuteScalarAsync_Count_ShouldReturnSeededRows() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            Dim count = Await connection.ExecuteScalarAsync(Of Integer)("SELECT COUNT(*) FROM clients")
            count.Should().Be(3)
        End Using
    End Function

    <SkippableFact>
    Public Sub ExecuteScalar_WithParameter_ShouldFilter()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            Dim activeCount = connection.ExecuteScalar(Of Integer)(
                "SELECT COUNT(*) FROM clients WHERE actif = ?", True)
            activeCount.Should().Be(2)
        End Using
    End Sub

    <SkippableFact>
    Public Async Function ExecuteNonQueryAsync_InsertData_ShouldReturnAffectedRows() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim table = NewTable()

        Using connection = _fixture.CreateConnection()
            Try
                Await connection.ExecuteNonQueryAsync($"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, nom TEXT(100))")
                Dim affected = Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (nom) VALUES (?)", {"Test"})
                affected.Should().Be(1)
            Finally
                connection.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Function

    <SkippableFact>
    Public Async Function ExecuteQueryAsync_ShouldReturnDataTable() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            Dim result = Await connection.ExecuteQueryAsync("SELECT nom FROM clients ORDER BY nom")
            result.Rows.Count.Should().Be(3)
            CStr(result.Rows(0)("nom")).Should().Be("Alice")
            CStr(result.Rows(1)("nom")).Should().Be("Bob")
            CStr(result.Rows(2)("nom")).Should().Be("Charlie")
        End Using
    End Function

    <SkippableFact>
    Public Sub ExecuteReader_ShouldReadRows()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            connection.Open()
            Using reader = connection.ExecuteReader("SELECT nom FROM clients WHERE actif = ? ORDER BY nom", True)
                reader.Read().Should().BeTrue()
                reader.GetString(0).Should().Be("Alice")
                reader.Read().Should().BeTrue()
                reader.GetString(0).Should().Be("Charlie")
                reader.Read().Should().BeFalse()
            End Using
        End Using
    End Sub

    <SkippableFact>
    Public Async Function Insert_MultipleTypedParameters_ShouldRoundTrip() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim table = NewTable()

        Using connection = _fixture.CreateConnection()
            Try
                Await connection.ExecuteNonQueryAsync(
                    $"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, nom TEXT(100), age LONG, salaire CURRENCY, embauche DATETIME)")

                Await connection.ExecuteNonQueryAsync(
                    $"INSERT INTO {table} (nom, age, salaire, embauche) VALUES (?, ?, ?, ?)",
                    {"Jean Dupont", 35, 45000.5D, New DateTime(2020, 1, 15)})

                Dim data = Await connection.ExecuteQueryAsync($"SELECT * FROM {table} WHERE nom = ?", {"Jean Dupont"})
                data.Rows.Count.Should().Be(1)
                Convert.ToInt32(data.Rows(0)("age")).Should().Be(35)
                Convert.ToDecimal(data.Rows(0)("salaire")).Should().Be(45000.5D)
            Finally
                connection.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Function

    <SkippableFact>
    Public Async Function Insert_ThenIdentity_ShouldReturnNewId() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim table = NewTable()

        Using connection = _fixture.CreateConnection()
            Try
                Await connection.ExecuteNonQueryAsync($"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, nom TEXT(100))")
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (nom) VALUES (?)", {"Premier"})
                Dim newId = Await connection.ExecuteScalarAsync(Of Integer)("SELECT @@IDENTITY")
                newId.Should().BeGreaterThan(0)
            Finally
                connection.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Function

    <SkippableFact>
    Public Async Function Transaction_Commit_ShouldPersistChanges() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim table = NewTable()

        Using connection = _fixture.CreateConnection()
            Try
                Await connection.ExecuteNonQueryAsync($"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, valeur LONG)")

                Await connection.BeginTransactionAsync()
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (valeur) VALUES (?)", {100})
                Await connection.CommitAsync()

                Dim count = Await connection.ExecuteScalarAsync(Of Integer)($"SELECT COUNT(*) FROM {table}")
                count.Should().Be(1)
            Finally
                connection.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Function

    <SkippableFact>
    Public Async Function Transaction_Rollback_ShouldRevertChanges() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim table = NewTable()

        Using connection = _fixture.CreateConnection()
            Try
                Await connection.ExecuteNonQueryAsync($"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, valeur LONG)")
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (valeur) VALUES (?)", {50})

                Await connection.BeginTransactionAsync()
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (valeur) VALUES (?)", {100})
                connection.IsTransactionActive.Should().BeTrue()
                Await connection.RollbackAsync()

                connection.IsTransactionActive.Should().BeFalse()
                Dim count = Await connection.ExecuteScalarAsync(Of Integer)($"SELECT COUNT(*) FROM {table}")
                count.Should().Be(1)
            Finally
                connection.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Function

    <SkippableFact>
    Public Async Function Update_And_Delete_WithParameters_ShouldAffectRows() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Dim table = NewTable()

        Using connection = _fixture.CreateConnection()
            Try
                Await connection.ExecuteNonQueryAsync($"CREATE TABLE {table} (id AUTOINCREMENT PRIMARY KEY, nom TEXT(50), statut TEXT(20))")
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (nom, statut) VALUES (?, ?)", {"A", "inactif"})
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (nom, statut) VALUES (?, ?)", {"B", "inactif"})
                Await connection.ExecuteNonQueryAsync($"INSERT INTO {table} (nom, statut) VALUES (?, ?)", {"C", "actif"})

                Dim updated = Await connection.ExecuteNonQueryAsync(
                    $"UPDATE {table} SET statut = ? WHERE statut = ?", {"actif", "inactif"})
                updated.Should().Be(2)

                Dim deleted = Await connection.ExecuteNonQueryAsync($"DELETE FROM {table} WHERE nom = ?", {"A"})
                deleted.Should().Be(1)

                Dim remaining = Await connection.ExecuteScalarAsync(Of Integer)($"SELECT COUNT(*) FROM {table}")
                remaining.Should().Be(2)
            Finally
                connection.ExecuteNonQuery($"DROP TABLE {table}")
            End Try
        End Using
    End Function

    <SkippableFact>
    Public Async Function ActionHistory_ShouldTrackActions() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            Await connection.OpenAsync()
            Await connection.CloseAsync()

            connection.LastAction.Should().Contain("CloseAsync")
            connection.ActionHistory.Should().NotBeEmpty()
        End Using
    End Function

    <SkippableFact>
    Public Sub Connections_ShouldHaveUniqueIds()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using c1 = _fixture.CreateConnection(), c2 = _fixture.CreateConnection()
            c1.Id.Should().NotBe(c2.Id)
        End Using
    End Sub

    <SkippableFact>
    Public Async Function DisposeAsync_ShouldCloseConnection() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim connection = _fixture.CreateConnection()
        Await connection.OpenAsync()
        Await connection.DisposeAsync()

        connection.State.Should().Be(ConnectionState.Closed)
    End Function

    <SkippableFact>
    Public Sub EngineVersion_And_ProviderName_ShouldBeAvailableWithoutOpening()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            ' Disponible sans ouvrir la connexion (lecture registre / nom du provider).
            connection.ProviderName.Should().Be("Microsoft.ACE.OLEDB.16.0")
            connection.EngineVersion.Should().StartWith("16.")
        End Using
    End Sub

End Class
