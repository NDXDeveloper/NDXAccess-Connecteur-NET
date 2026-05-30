Imports System.Data.OleDb
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices

Namespace NDXAccess

    ''' <summary>
    ''' Création de fichiers de base Access (.accdb) via ADOX (liaison tardive), sans
    ''' dépendance d'assembly d'interopérabilité.
    ''' </summary>
    Friend Module AccessDatabaseManager

        Friend Sub Create(databasePath As String, password As String, provider As String)
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("Le chemin de la base à créer est requis.", NameOf(databasePath))
            End If
            If File.Exists(databasePath) Then
                Throw New IOException($"Le fichier existe déjà : {databasePath}")
            End If
            If String.IsNullOrWhiteSpace(provider) Then
                provider = AccessConnectionOptions.DefaultProvider
            End If

            ' Vérifie la cohérence x86/x64 avec un message clair.
            AccessProviderHelper.EnsureProviderAvailable(provider)

            Dim builder As New OleDbConnectionStringBuilder() With {.Provider = provider}
            builder("Data Source") = databasePath
            If Not String.IsNullOrEmpty(password) Then
                builder("Jet OLEDB:Database Password") = password
            End If

            Dim catalogType = Type.GetTypeFromProgID("ADOX.Catalog")
            If catalogType Is Nothing Then
                Throw New AccessConnectionException(
                    "ADOX.Catalog est introuvable : impossible de créer la base. " &
                    "Installez Microsoft Access Database Engine dans l'architecture du processus.")
            End If

            Dim catalog As Object = Nothing
            Try
                catalog = Activator.CreateInstance(catalogType)
                catalogType.InvokeMember("Create", BindingFlags.InvokeMethod, Nothing, catalog, New Object() {builder.ConnectionString})
            Catch ex As TargetInvocationException When ex.InnerException IsNot Nothing
                Throw New AccessConnectionException(
                    "Échec de la création de la base Access : " & ex.InnerException.Message, ex.InnerException)
            Finally
                If catalog IsNot Nothing AndAlso Marshal.IsComObject(catalog) Then
                    Marshal.FinalReleaseComObject(catalog)
                End If
            End Try
        End Sub

    End Module

End Namespace
