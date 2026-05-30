Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>Tests unitaires des informations de version de la bibliothèque (aucune base requise).</summary>
<Trait("Category", "Unit")>
Public Class NDXAccessInfoTests

    <Fact>
    Public Sub Version_ShouldBeAvailable()
        NDXAccessInfo.Version.Should().NotBeNullOrEmpty()
        NDXAccessInfo.Version.Should().StartWith("1.")
    End Sub

    <Fact>
    Public Sub InformationalVersion_ShouldStartWithMajorVersion()
        NDXAccessInfo.InformationalVersion.Should().StartWith("1.")
    End Sub

    <Fact>
    Public Sub ProductName_ShouldBeNDXAccess()
        NDXAccessInfo.ProductName.Should().Be("NDXAccess")
    End Sub

    <Fact>
    Public Sub AccessConnection_Version_ShouldMatchInformationalVersion()
        AccessConnection.Version.Should().Be(NDXAccessInfo.InformationalVersion)
    End Sub

End Class
