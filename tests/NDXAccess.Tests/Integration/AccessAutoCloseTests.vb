Imports System.Data
Imports System.Threading
Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>
''' Tests d'intégration du timer de fermeture automatique des connexions inactives.
''' (Légèrement sensibles au timing : marges d'attente généreuses.)
''' </summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class AccessAutoCloseTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    <SkippableFact>
    Public Sub IdleConnection_ShouldAutoCloseAfterTimeout()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim options = _fixture.GetOptions()
        options.AutoCloseTimeoutMs = 300
        options.IsPrimaryConnection = False

        Using connection As IAccessConnection = New AccessConnection(options)
            connection.Open()
            connection.State.Should().Be(ConnectionState.Open)

            Thread.Sleep(1500)   ' >> 300 ms : le timer doit avoir fermé la connexion

            connection.State.Should().Be(ConnectionState.Closed)
        End Using
    End Sub

    <SkippableFact>
    Public Sub PrimaryConnection_ShouldNotAutoClose()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim options = _fixture.GetOptions()
        options.AutoCloseTimeoutMs = 300
        options.IsPrimaryConnection = True   ' connexion principale : pas de fermeture auto

        Using connection As IAccessConnection = New AccessConnection(options)
            connection.Open()
            Thread.Sleep(1500)
            connection.State.Should().Be(ConnectionState.Open)
        End Using
    End Sub

End Class
