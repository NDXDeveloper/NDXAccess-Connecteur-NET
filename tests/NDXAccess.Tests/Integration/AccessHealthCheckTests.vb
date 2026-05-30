Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>Tests d'intégration de AccessHealthCheck (ignorés si ACE absent).</summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class AccessHealthCheckTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    <SkippableFact>
    Public Async Function CheckHealthAsync_WhenHealthy_ShouldReturnHealthyResult() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim healthCheck = New AccessHealthCheck(_fixture.CreateFactory())
        Dim result = Await healthCheck.CheckHealthAsync()

        result.Should().NotBeNull()
        result.IsHealthy.Should().BeTrue()
        result.Message.Should().Contain("fonctionnelle")
        result.Exception.Should().BeNull()
        result.DatabaseInfo.Should().NotBeNull()
    End Function

    <SkippableFact>
    Public Async Function CheckHealthAsync_WhenFileMissing_ShouldReturnUnhealthyResult() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim badOptions = New AccessConnectionOptions With {
            .DatabasePath = IO.Path.Combine(IO.Path.GetTempPath(), $"inexistant_{Guid.NewGuid():N}.accdb"),
            .Provider = AccessDatabaseFixture.Provider
        }
        Dim healthCheck = New AccessHealthCheck(New AccessConnectionFactory(badOptions))

        Dim result = Await healthCheck.CheckHealthAsync()

        result.IsHealthy.Should().BeFalse()
        result.Message.Should().Contain("Erreur")
        result.Exception.Should().NotBeNull()
    End Function

    <SkippableFact>
    Public Async Function GetDatabaseInfoAsync_ShouldReturnFileAndSchemaInfo() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim healthCheck = New AccessHealthCheck(_fixture.CreateFactory())
        Dim info = Await healthCheck.GetDatabaseInfoAsync()

        info.Should().NotBeNull()
        info.Provider.Should().Contain("ACE.OLEDB")
        info.FilePath.Should().Be(_fixture.DatabasePath)
        info.FileSizeBytes.Should().BeGreaterThan(0)
        info.MaxSizeBytes.Should().Be(2_147_483_648L)
        info.UsagePercent.Should().BeLessThan(100.0)
        info.IsApproachingSizeLimit.Should().BeFalse()
        info.UserTableCount.Should().BeGreaterThanOrEqualTo(1)
    End Function

    <SkippableFact>
    Public Async Function GetDatabaseInfoAsync_ShouldExposeEngineAndFileFormatVersions() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim healthCheck = New AccessHealthCheck(_fixture.CreateFactory())
        Dim info = Await healthCheck.GetDatabaseInfoAsync()

        ' Version du moteur ACE (DLL ou génération) : commence par "16.".
        info.EngineVersion.Should().StartWith("16.")
        ' Version du format de fichier (Jet/ACE), distincte du moteur.
        info.FileFormatVersion.Should().NotBeNullOrEmpty()
    End Function

End Class
