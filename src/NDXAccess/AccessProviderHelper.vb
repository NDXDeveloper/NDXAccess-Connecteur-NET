Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.Win32

Namespace NDXAccess

    ''' <summary>
    ''' Utilitaires de détection des providers OLE DB ACE installés et de leur
    ''' compatibilité avec l'architecture (x86/x64) du processus courant.
    ''' </summary>
    ''' <remarks>
    ''' Le piège classique d'Access : le provider ACE installé doit correspondre au
    ''' bitness de l'application. Un ACE x64 est invisible depuis un processus 32 bits
    ''' et inversement. <see cref="OleDbEnumerator"/> n'énumère que les providers
    ''' visibles pour l'architecture courante : c'est donc un test fiable du mismatch.
    ''' </remarks>
    Public Module AccessProviderHelper

        ''' <summary>
        ''' Architecture du processus courant : "x64" ou "x86".
        ''' </summary>
        Public ReadOnly Property CurrentProcessArchitecture As String
            Get
                Return If(Environment.Is64BitProcess, "x64", "x86")
            End Get
        End Property

        ''' <summary>
        ''' Retourne la liste des providers OLE DB visibles pour l'architecture courante.
        ''' </summary>
        Public Function GetAvailableProviders() As IReadOnlyList(Of String)
            Dim result As New List(Of String)()

            Try
                Dim enumerator As New OleDbEnumerator()
                Using table As DataTable = enumerator.GetElements()
                    If table.Columns.Contains("SOURCES_NAME") Then
                        For Each row As DataRow In table.Rows
                            Dim name = TryCast(row("SOURCES_NAME"), String)
                            If Not String.IsNullOrWhiteSpace(name) Then
                                result.Add(name)
                            End If
                        Next
                    End If
                End Using
            Catch
                ' L'énumération peut échouer sur certaines configurations : on retourne ce qu'on a.
            End Try

            Return result
        End Function

        ''' <summary>
        ''' Indique si le provider spécifié est disponible pour l'architecture courante.
        ''' </summary>
        ''' <param name="providerName">Ex. "Microsoft.ACE.OLEDB.16.0".</param>
        Public Function IsProviderAvailable(providerName As String) As Boolean
            If String.IsNullOrWhiteSpace(providerName) Then
                Return False
            End If

            For Each name In GetAvailableProviders()
                If String.Equals(name, providerName, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next

            Return False
        End Function

        ''' <summary>
        ''' Vérifie que le provider est disponible et lève une exception explicite sinon,
        ''' en distinguant le cas "non installé" du cas "mauvaise architecture".
        ''' </summary>
        ''' <exception cref="AccessProviderNotFoundException">Provider absent pour l'architecture courante.</exception>
        Public Sub EnsureProviderAvailable(providerName As String)
            If IsProviderAvailable(providerName) Then
                Return
            End If

            Dim arch = CurrentProcessArchitecture
            Dim available = GetAvailableProviders()
            Dim aceProviders = available.Where(Function(p) p.StartsWith("Microsoft.ACE.OLEDB", StringComparison.OrdinalIgnoreCase) OrElse
                                                            p.StartsWith("Microsoft.Jet.OLEDB", StringComparison.OrdinalIgnoreCase)).ToList()

            Dim sb As New Text.StringBuilder()
            sb.AppendLine($"Le provider OLE DB '{providerName}' n'est pas disponible pour un processus {arch}.")

            If aceProviders.Count > 0 Then
                sb.AppendLine($"Des providers Access sont installés mais dans une autre architecture ({String.Join(", ", aceProviders)}).")
                sb.AppendLine($"Compilez l'application en {If(arch = "x64", "x86", "x64")} OU installez l'Access Database Engine en {arch}.")
            Else
                sb.AppendLine("Aucun provider Microsoft ACE/Jet n'est installé.")
                sb.AppendLine($"Installez 'Microsoft Access Database Engine 2016 Redistributable' en version {arch} :")
                sb.AppendLine("https://www.microsoft.com/download/details.aspx?id=54920")
            End If

            sb.Append("Note : pour installer ACE x86 ET x64 côte à côte, utilisez l'option /quiet /passive du redistribuable.")

            Throw New AccessProviderNotFoundException(providerName, arch, sb.ToString())
        End Sub

        ''' <summary>
        ''' Retourne la version du moteur ACE pour le provider donné : la version exacte du
        ''' DLL (ex. "16.0.5011.1000", lue dans le registre) si disponible, sinon la
        ''' génération du provider extraite de son nom (ex. "16.0").
        ''' </summary>
        ''' <remarks>
        ''' À ne pas confondre avec la version du FORMAT de fichier (Jet 4 = "04.00.0000",
        ''' exposée par <see cref="DatabaseInfo.FileFormatVersion"/>).
        ''' </remarks>
        Public Function GetEngineVersion(Optional providerName As String = AccessConnectionOptions.DefaultProvider) As String
            If String.IsNullOrWhiteSpace(providerName) Then
                providerName = AccessConnectionOptions.DefaultProvider
            End If

            ' 1) Version exacte du DLL ACE via le registre (s'adapte au bitness du processus).
            Try
                Dim dllPath = GetProviderDllPath(providerName)
                If Not String.IsNullOrEmpty(dllPath) AndAlso File.Exists(dllPath) Then
                    Dim fileVersion = FileVersionInfo.GetVersionInfo(dllPath).FileVersion
                    If Not String.IsNullOrWhiteSpace(fileVersion) Then
                        Return fileVersion
                    End If
                End If
            Catch
                ' Registre indisponible / accès refusé : on bascule sur le repli.
            End Try

            ' 2) Repli : génération du provider (ex. "16.0").
            Return GetProviderGeneration(providerName)
        End Function

        ''' <summary>
        ''' Extrait la génération d'un nom de provider : "Microsoft.ACE.OLEDB.16.0" -> "16.0".
        ''' </summary>
        Public Function GetProviderGeneration(providerName As String) As String
            If String.IsNullOrEmpty(providerName) Then Return Nothing
            Dim match = Regex.Match(providerName, "(\d+\.\d+)")
            Return If(match.Success, match.Value, Nothing)
        End Function

        ''' <summary>Chemin du DLL du provider (via HKCR\&lt;progid&gt;\CLSID\...\InprocServer32).</summary>
        Private Function GetProviderDllPath(providerName As String) As String
            Using clsidKey = Registry.ClassesRoot.OpenSubKey($"{providerName}\CLSID")
                If clsidKey Is Nothing Then Return Nothing
                Dim clsid = TryCast(clsidKey.GetValue(Nothing), String)
                If String.IsNullOrEmpty(clsid) Then Return Nothing

                Using inproc = Registry.ClassesRoot.OpenSubKey($"CLSID\{clsid}\InprocServer32")
                    If inproc Is Nothing Then Return Nothing
                    Return TryCast(inproc.GetValue(Nothing), String)
                End Using
            End Using
        End Function

    End Module

End Namespace
