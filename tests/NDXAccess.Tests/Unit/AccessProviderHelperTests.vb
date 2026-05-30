Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>
''' Tests unitaires de la détection x86/x64 du provider (aucune base requise).
''' </summary>
<Trait("Category", "Unit")>
Public Class AccessProviderHelperTests

    <Fact>
    Public Sub CurrentProcessArchitecture_ShouldBeX86OrX64()
        AccessProviderHelper.CurrentProcessArchitecture.Should().BeOneOf("x86", "x64")
    End Sub

    <Fact>
    Public Sub GetAvailableProviders_ShouldNeverReturnNull()
        AccessProviderHelper.GetAvailableProviders().Should().NotBeNull()
    End Sub

    <Fact>
    Public Sub IsProviderAvailable_WithUnknownProvider_ShouldReturnFalse()
        AccessProviderHelper.IsProviderAvailable("Provider.Inexistant.99.0").Should().BeFalse()
    End Sub

    <Fact>
    Public Sub IsProviderAvailable_WithNullOrEmpty_ShouldReturnFalse()
        AccessProviderHelper.IsProviderAvailable(Nothing).Should().BeFalse()
        AccessProviderHelper.IsProviderAvailable("").Should().BeFalse()
    End Sub

    <Fact>
    Public Sub EnsureProviderAvailable_WithUnknownProvider_ShouldThrowDescriptiveException()
        Dim act As Action = Sub() AccessProviderHelper.EnsureProviderAvailable("Provider.Inexistant.99.0")

        Dim assertion = act.Should().Throw(Of AccessProviderNotFoundException)()
        assertion.Which.ProviderName.Should().Be("Provider.Inexistant.99.0")
        assertion.Which.ProcessArchitecture.Should().BeOneOf("x86", "x64")
        assertion.Which.Message.Should().Contain(AccessProviderHelper.CurrentProcessArchitecture)
    End Sub

    <Theory>
    <InlineData("Microsoft.ACE.OLEDB.16.0", "16.0")>
    <InlineData("Microsoft.ACE.OLEDB.12.0", "12.0")>
    <InlineData("Microsoft.Jet.OLEDB.4.0", "4.0")>
    Public Sub GetProviderGeneration_ShouldParseVersionFromName(providerName As String, expected As String)
        AccessProviderHelper.GetProviderGeneration(providerName).Should().Be(expected)
    End Sub

    <Fact>
    Public Sub GetEngineVersion_ShouldReturnAce16VersionOrGeneration()
        ' Avec ACE 16 : version exacte du DLL (ex. "16.0.5011.1000"). Sans ACE : repli "16.0".
        ' Dans les deux cas, commence par "16".
        Dim version = AccessProviderHelper.GetEngineVersion("Microsoft.ACE.OLEDB.16.0")
        version.Should().NotBeNullOrEmpty()
        version.Should().StartWith("16.")
    End Sub

End Class
