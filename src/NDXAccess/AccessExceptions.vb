Namespace NDXAccess

    ''' <summary>
    ''' Exception de base pour les erreurs du connecteur NDXAccess.
    ''' </summary>
    Public Class AccessConnectionException
        Inherits Exception

        ''' <summary>Crée l'exception avec un message.</summary>
        Public Sub New(message As String)
            MyBase.New(message)
        End Sub

        ''' <summary>Crée l'exception avec un message et une exception interne.</summary>
        Public Sub New(message As String, innerException As Exception)
            MyBase.New(message, innerException)
        End Sub
    End Class

    ''' <summary>
    ''' Levée lorsque le provider OLE DB ACE demandé n'est pas disponible
    ''' dans l'architecture (x86/x64) du processus courant.
    ''' </summary>
    ''' <remarks>
    ''' Le provider ACE installé doit correspondre à l'architecture de l'application :
    ''' un ACE x64 n'est pas visible depuis un processus x86 et inversement.
    ''' </remarks>
    Public NotInheritable Class AccessProviderNotFoundException
        Inherits AccessConnectionException

        ''' <summary>Nom du provider recherché.</summary>
        Public ReadOnly Property ProviderName As String

        ''' <summary>Architecture du processus courant ("x86" ou "x64").</summary>
        Public ReadOnly Property ProcessArchitecture As String

        Public Sub New(providerName As String, processArchitecture As String, message As String)
            MyBase.New(message)
            Me.ProviderName = providerName
            Me.ProcessArchitecture = processArchitecture
        End Sub
    End Class

    ''' <summary>
    ''' Encapsule une <see cref="System.Data.OleDb.OleDbException"/> en y ajoutant un
    ''' message clair et le code d'erreur natif ACE/Jet (<see cref="NativeError"/>).
    ''' L'exception OLE DB d'origine reste disponible via <see cref="Exception.InnerException"/>.
    ''' </summary>
    Public NotInheritable Class AccessQueryException
        Inherits AccessConnectionException

        ''' <summary>Code d'erreur natif du moteur ACE/Jet (0 si inconnu).</summary>
        Public ReadOnly Property NativeError As Integer

        ''' <summary>Indique si l'erreur d'origine est de nature transitoire (verrou, contention).</summary>
        Public ReadOnly Property IsTransient As Boolean

        Public Sub New(message As String, nativeError As Integer, isTransient As Boolean, innerException As Exception)
            MyBase.New(message, innerException)
            Me.NativeError = nativeError
            Me.IsTransient = isTransient
        End Sub
    End Class

End Namespace
