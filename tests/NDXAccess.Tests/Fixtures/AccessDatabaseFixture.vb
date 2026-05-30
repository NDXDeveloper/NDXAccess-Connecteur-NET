Imports System.Data.OleDb
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Extensions.Logging
Imports NDXAccess
Imports Xunit

''' <summary>
''' Fixture des tests d'intégration : crée une base Access (.accdb) temporaire via ADOX
''' (liaison tardive) si le provider ACE est disponible pour l'architecture courante.
''' Si ACE est absent, <see cref="ProviderAvailable"/> vaut False et les tests
''' d'intégration sont ignorés (Skip) plutôt qu'en échec.
''' </summary>
Public NotInheritable Class AccessDatabaseFixture
    Implements IAsyncLifetime

    Public Const Provider As String = "Microsoft.ACE.OLEDB.16.0"

    Public ReadOnly Property DatabasePath As String
    Public ReadOnly Property ProviderAvailable As Boolean
    Public ReadOnly Property SkipReason As String

    Public Sub New()
        _DatabasePath = Path.Combine(Path.GetTempPath(), $"ndxaccess_test_{Guid.NewGuid():N}.accdb")
        _ProviderAvailable = AccessProviderHelper.IsProviderAvailable(Provider)
        If Not _ProviderAvailable Then
            _SkipReason = $"Provider '{Provider}' indisponible pour un processus {AccessProviderHelper.CurrentProcessArchitecture}. " &
                          "Installez Microsoft Access Database Engine 2016 dans cette architecture."
        End If
    End Sub

    Public Function GetOptions() As AccessConnectionOptions
        Return New AccessConnectionOptions With {
            .DatabasePath = DatabasePath,
            .Provider = Provider
        }
    End Function

    Public Function CreateFactory(Optional loggerFactory As ILoggerFactory = Nothing) As IAccessConnectionFactory
        Return New AccessConnectionFactory(GetOptions(), loggerFactory)
    End Function

    Public Function CreateConnection(Optional logger As ILogger(Of AccessConnection) = Nothing) As IAccessConnection
        Return New AccessConnection(GetOptions(), logger)
    End Function

    Public Function InitializeAsync() As Task Implements IAsyncLifetime.InitializeAsync
        If Not ProviderAvailable Then
            Return Task.CompletedTask
        End If

        CreateEmptyDatabase(DatabasePath)
        SeedSchema()
        Return Task.CompletedTask
    End Function

    Public Function DisposeAsync() As Task Implements IAsyncLifetime.DisposeAsync
        Try
            OleDbConnection.ReleaseObjectPool()
        Catch
        End Try
        Try
            If File.Exists(DatabasePath) Then File.Delete(DatabasePath)
        Catch
            ' Le verrou .laccdb peut subsister brièvement : suppression best-effort.
        End Try
        Return Task.CompletedTask
    End Function

    ''' <summary>Crée un fichier .accdb vide via ADOX.Catalog.Create (late binding).</summary>
    Private Shared Sub CreateEmptyDatabase(path As String)
        If File.Exists(path) Then File.Delete(path)

        Dim connectionString = $"Provider={Provider};Data Source={path};"
        Dim catalogType = Type.GetTypeFromProgID("ADOX.Catalog")
        If catalogType Is Nothing Then
            Throw New InvalidOperationException("ADOX.Catalog introuvable : impossible de créer la base de test.")
        End If

        Dim catalog As Object = Nothing
        Try
            catalog = Activator.CreateInstance(catalogType)
            catalogType.InvokeMember("Create", BindingFlags.InvokeMethod, Nothing, catalog, New Object() {connectionString})
        Finally
            If catalog IsNot Nothing AndAlso Marshal.IsComObject(catalog) Then
                Marshal.FinalReleaseComObject(catalog)
            End If
        End Try
    End Sub

    ''' <summary>Crée le schéma de base partagé par les tests.</summary>
    Private Sub SeedSchema()
        Using connection = CreateConnection()
            connection.Open()
            connection.ExecuteNonQuery(
                "CREATE TABLE clients (id AUTOINCREMENT PRIMARY KEY, nom TEXT(100), email TEXT(200), actif YESNO)")
            connection.ExecuteNonQuery("INSERT INTO clients (nom, email, actif) VALUES (?, ?, ?)", "Alice", "alice@example.com", True)
            connection.ExecuteNonQuery("INSERT INTO clients (nom, email, actif) VALUES (?, ?, ?)", "Bob", "bob@example.com", False)
            connection.ExecuteNonQuery("INSERT INTO clients (nom, email, actif) VALUES (?, ?, ?)", "Charlie", "charlie@example.com", True)
        End Using
    End Sub

End Class

''' <summary>Collection partageant la fixture base Access (exécution séquentielle).</summary>
<CollectionDefinition("Access")>
Public Class AccessCollection
    Implements ICollectionFixture(Of AccessDatabaseFixture)
End Class
