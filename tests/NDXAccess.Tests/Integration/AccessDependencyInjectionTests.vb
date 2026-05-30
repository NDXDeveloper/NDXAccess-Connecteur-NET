Imports FluentAssertions
Imports Microsoft.Extensions.DependencyInjection
Imports NDXAccess
Imports NDXAccess.Extensions
Imports Xunit

''' <summary>
''' Tests d'intégration de l'injection de dépendances (AddNDXAccess) : résolution et
''' utilisation réelle de la factory, de la connexion et du health check.
''' </summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class AccessDependencyInjectionTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    <SkippableFact>
    Public Async Function AddNDXAccess_ShouldRegisterAndResolveServices() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim services As New ServiceCollection()
        services.AddNDXAccess(Sub(options) options.DatabasePath = _fixture.DatabasePath)

        Using provider = services.BuildServiceProvider()
            ' Factory -> connexion -> requête
            Dim factory = provider.GetRequiredService(Of IAccessConnectionFactory)()
            factory.Should().NotBeNull()
            Using conn = factory.CreateConnection()
                Dim count = conn.ExecuteScalar(Of Integer)("SELECT COUNT(*) FROM clients")
                count.Should().Be(3)
            End Using

            ' IAccessConnection résolu directement (transient)
            Dim resolved = provider.GetRequiredService(Of IAccessConnection)()
            resolved.Should().NotBeNull()
            resolved.Dispose()

            ' Health check résolu
            Dim healthCheck = provider.GetRequiredService(Of AccessHealthCheck)()
            Dim result = Await healthCheck.CheckHealthAsync()
            result.IsHealthy.Should().BeTrue()
        End Using
    End Function

End Class
