Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Threading
Imports FluentAssertions
Imports NDXAccess
Imports Xunit

''' <summary>
''' Tests d'intégration de la résilience : retry sur verrou transitoire (ouverture
''' exclusive concurrente) et prise en compte du CancellationToken.
''' </summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class AccessResilienceTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    <SkippableFact>
    Public Sub Retry_OnExclusiveLock_ShouldRetryThenThrowTransient()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Dim exclusive = _fixture.GetOptions()
        exclusive.OpenExclusive = True

        Using connA As IAccessConnection = New AccessConnection(exclusive)
            connA.Open()   ' verrou exclusif détenu pendant tout le test
            Try
                ' --- Sans retry : échoue rapidement ---
                Dim noRetry = _fixture.GetOptions()
                noRetry.EnableRetryOnTransientErrors = False
                Dim swNo As New Stopwatch()
                swNo.Start()
                Dim actNo As Action =
                    Sub()
                        Using b As IAccessConnection = New AccessConnection(noRetry)
                            b.Open()
                        End Using
                    End Sub
                actNo.Should().Throw(Of AccessQueryException)()
                swNo.Stop()

                ' --- Avec retry : back-off 150+300+600 ms avant l'échec ---
                Dim withRetry = _fixture.GetOptions()
                withRetry.MaxRetries = 3
                withRetry.RetryBaseDelayMs = 150
                Dim swYes As New Stopwatch()
                swYes.Start()
                Dim actYes As Action =
                    Sub()
                        Using b As IAccessConnection = New AccessConnection(withRetry)
                            b.Open()
                        End Using
                    End Sub
                Dim ex = actYes.Should().Throw(Of AccessQueryException)().Which
                swYes.Stop()

                ex.IsTransient.Should().BeTrue()
                ' Le chemin avec retry prend nettement plus longtemps (boucle de back-off).
                swYes.ElapsedMilliseconds.Should().BeGreaterThan(swNo.ElapsedMilliseconds + 300)
            Finally
                ' Libère le pool pour relâcher le verrou exclusif avant les tests suivants.
                OleDbConnection.ReleaseObjectPool()
            End Try
        End Using

        OleDbConnection.ReleaseObjectPool()
    End Sub

    <SkippableFact>
    Public Async Function ExecuteAsync_WithCancelledToken_ShouldThrowOperationCanceled() As Task
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)

        Using connection = _fixture.CreateConnection()
            Using cts As New CancellationTokenSource()
                cts.Cancel()
                Dim act As Func(Of Task) =
                    Function() connection.ExecuteScalarAsync(Of Integer)("SELECT COUNT(*) FROM clients", cancellationToken:=cts.Token)
                Await act.Should().ThrowAsync(Of OperationCanceledException)()
            End Using
        End Using
    End Function

End Class
