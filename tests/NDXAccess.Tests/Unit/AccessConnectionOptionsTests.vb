Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>Tests unitaires de AccessConnectionOptions (aucune base requise).</summary>
<Trait("Category", "Unit")>
Public Class AccessConnectionOptionsTests

    <Fact>
    Public Sub DefaultValues_ShouldBeCorrect()
        Dim options As New AccessConnectionOptions()

        options.DatabasePath.Should().BeEmpty()
        options.Password.Should().BeEmpty()
        options.Provider.Should().Be("Microsoft.ACE.OLEDB.16.0")
        options.ConnectionString.Should().BeNull()
        options.OpenExclusive.Should().BeFalse()
        options.OpenReadOnly.Should().BeFalse()
        options.PersistSecurityInfo.Should().BeFalse()
        options.IsPrimaryConnection.Should().BeFalse()
        options.AutoCloseTimeoutMs.Should().Be(60_000)
        options.DisableAutoClose.Should().BeFalse()
        options.ValidateProvider.Should().BeTrue()
    End Sub

    <Fact>
    Public Sub MaxDatabaseSize_ShouldBeTwoGigabytes()
        AccessConnectionOptions.MaxDatabaseSizeBytes.Should().Be(2_147_483_648L)
    End Sub

    <Fact>
    Public Sub ResilienceDefaults_ShouldBeReasonable()
        Dim options As New AccessConnectionOptions()
        options.EnableRetryOnTransientErrors.Should().BeTrue()
        options.MaxRetries.Should().Be(3)
        options.RetryBaseDelayMs.Should().Be(100)
        options.TranslateErrors.Should().BeTrue()
    End Sub

    <Fact>
    Public Sub BuildConnectionString_WithProperties_ShouldContainProviderAndDataSource()
        Dim options As New AccessConnectionOptions With {
            .DatabasePath = "C:\data\ma_base.accdb"
        }

        Dim result = options.BuildConnectionString()

        result.Should().Contain("Provider=Microsoft.ACE.OLEDB.16.0")
        result.Should().Contain("Data Source=C:\data\ma_base.accdb")
    End Sub

    <Fact>
    Public Sub BuildConnectionString_WithPassword_ShouldContainJetPassword()
        Dim options As New AccessConnectionOptions With {
            .DatabasePath = "C:\data\secure.accdb",
            .Password = "s3cret"
        }

        Dim result = options.BuildConnectionString()

        result.Should().Contain("Jet OLEDB:Database Password=s3cret")
    End Sub

    <Fact>
    Public Sub BuildConnectionString_WithoutPassword_ShouldNotContainJetPassword()
        Dim options As New AccessConnectionOptions With {
            .DatabasePath = "C:\data\open.accdb"
        }

        Dim result = options.BuildConnectionString()

        result.Should().NotContain("Jet OLEDB:Database Password")
    End Sub

    <Fact>
    Public Sub BuildConnectionString_Exclusive_ShouldSetShareExclusiveMode()
        Dim options As New AccessConnectionOptions With {
            .DatabasePath = "C:\data\ex.accdb",
            .OpenExclusive = True
        }

        options.BuildConnectionString().Should().Contain("Share Exclusive")
    End Sub

    <Fact>
    Public Sub BuildConnectionString_ReadOnly_ShouldSetReadMode()
        Dim options As New AccessConnectionOptions With {
            .DatabasePath = "C:\data\ro.accdb",
            .OpenReadOnly = True
        }

        options.BuildConnectionString().Should().Contain("Mode=Read")
    End Sub

    <Fact>
    Public Sub BuildConnectionString_WithConnectionString_ShouldOverrideProperties()
        Dim raw = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\autre.mdb;"
        Dim options As New AccessConnectionOptions With {
            .DatabasePath = "C:\ignored.accdb",
            .ConnectionString = raw
        }

        options.BuildConnectionString().Should().Be(raw)
    End Sub

    <Fact>
    Public Sub ResolveProviderName_FromConnectionString_ShouldReturnEmbeddedProvider()
        Dim options As New AccessConnectionOptions With {
            .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\x.mdb;"
        }

        options.ResolveProviderName().Should().Be("Microsoft.ACE.OLEDB.12.0")
    End Sub

    <Fact>
    Public Sub ResolveDatabasePath_FromConnectionString_ShouldReturnDataSource()
        Dim options As New AccessConnectionOptions With {
            .ConnectionString = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\base\reseau.accdb;"
        }

        options.ResolveDatabasePath().Should().Be("C:\base\reseau.accdb")
    End Sub

End Class
