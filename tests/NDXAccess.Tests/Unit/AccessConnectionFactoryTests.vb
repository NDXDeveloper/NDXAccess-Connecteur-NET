Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>
''' Tests unitaires de AccessConnectionFactory. On désactive la validation du provider
''' (ValidateProvider = False) pour que ces tests s'exécutent même sans ACE installé :
''' aucune connexion réelle n'est ouverte ici.
''' </summary>
<Trait("Category", "Unit")>
Public Class AccessConnectionFactoryTests

    Private Shared Function BuildOptions() As AccessConnectionOptions
        Return New AccessConnectionOptions With {
            .DatabasePath = "C:\data\test.accdb",
            .ValidateProvider = False
        }
    End Function

    <Fact>
    Public Sub Constructor_WithNullOptions_ShouldThrowArgumentNullException()
        Dim act As Action =
            Sub()
                Dim f = New AccessConnectionFactory(CType(Nothing, AccessConnectionOptions))
            End Sub
        act.Should().Throw(Of ArgumentNullException)()
    End Sub

    <Fact>
    Public Sub Constructor_WithDatabasePath_ShouldWork()
        Dim factory = New AccessConnectionFactory("C:\data\test.accdb")
        factory.Should().NotBeNull()
    End Sub

    <Fact>
    Public Sub CreateConnection_ShouldReturnNonPrimaryConnection()
        Dim factory = New AccessConnectionFactory(BuildOptions())

        Using connection = factory.CreateConnection()
            connection.Should().NotBeNull()
            connection.Should().BeOfType(Of AccessConnection)()
            connection.IsPrimaryConnection.Should().BeFalse()
        End Using
    End Sub

    <Fact>
    Public Sub CreatePrimaryConnection_ShouldReturnPrimaryConnection()
        Dim factory = New AccessConnectionFactory(BuildOptions())

        Using connection = factory.CreatePrimaryConnection()
            connection.IsPrimaryConnection.Should().BeTrue()
        End Using
    End Sub

    <Fact>
    Public Sub CreateConnection_WithConfigure_ShouldApplyConfiguration()
        Dim factory = New AccessConnectionFactory(BuildOptions())

        Using connection = factory.CreateConnection(
            Sub(opts)
                opts.IsPrimaryConnection = True
                opts.DisableAutoClose = True
            End Sub)
            connection.IsPrimaryConnection.Should().BeTrue()
        End Using
    End Sub

    <Fact>
    Public Sub CreateConnection_ShouldGenerateUniqueIds()
        Dim factory = New AccessConnectionFactory(BuildOptions())

        Using c1 = factory.CreateConnection(),
              c2 = factory.CreateConnection(),
              c3 = factory.CreateConnection()
            c1.Id.Should().NotBe(c2.Id)
            c2.Id.Should().NotBe(c3.Id)
            c1.Id.Should().NotBe(c3.Id)
        End Using
    End Sub

    <Fact>
    Public Sub CreateConnection_ShouldNotShareTransactionState()
        Dim factory = New AccessConnectionFactory(BuildOptions())

        Using connection = factory.CreateConnection()
            connection.IsTransactionActive.Should().BeFalse()
            connection.State.Should().Be(System.Data.ConnectionState.Closed)
        End Using
    End Sub

End Class
