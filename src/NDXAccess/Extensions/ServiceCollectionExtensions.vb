Imports System.Runtime.CompilerServices
Imports Microsoft.Extensions.DependencyInjection
Imports Microsoft.Extensions.DependencyInjection.Extensions
Imports NDXAccess

Namespace NDXAccess.Extensions

    ''' <summary>
    ''' Extensions d'injection de dépendances pour Microsoft.Extensions.DependencyInjection.
    ''' </summary>
    Public Module ServiceCollectionExtensions

        ''' <summary>
        ''' Enregistre les services NDXAccess dans le conteneur d'injection de dépendances.
        ''' </summary>
        ''' <param name="services">Collection de services.</param>
        ''' <param name="configure">Action de configuration des options.</param>
        <Extension()>
        Public Function AddNDXAccess(services As IServiceCollection, configure As Action(Of AccessConnectionOptions)) As IServiceCollection
            ArgumentNullException.ThrowIfNull(services)
            ArgumentNullException.ThrowIfNull(configure)

            Dim options As New AccessConnectionOptions()
            configure(options)

            services.TryAddSingleton(options)
            services.TryAddSingleton(Of IAccessConnectionFactory, AccessConnectionFactory)()
            services.TryAddTransient(Of IAccessConnection)(
                Function(sp)
                    Dim factory = sp.GetRequiredService(Of IAccessConnectionFactory)()
                    Return factory.CreateConnection()
                End Function)
            services.TryAddSingleton(Of AccessHealthCheck)()

            Return services
        End Function

        ''' <summary>
        ''' Enregistre les services NDXAccess à partir d'un chemin de fichier de base.
        ''' </summary>
        ''' <param name="services">Collection de services.</param>
        ''' <param name="databasePath">Chemin du fichier .accdb / .mdb.</param>
        ''' <param name="password">Mot de passe optionnel.</param>
        <Extension()>
        Public Function AddNDXAccess(services As IServiceCollection, databasePath As String, Optional password As String = Nothing) As IServiceCollection
            Return services.AddNDXAccess(
                Sub(options)
                    options.DatabasePath = databasePath
                    options.Password = If(password, String.Empty)
                End Sub)
        End Function

    End Module

End Namespace
