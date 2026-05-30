Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices

Namespace NDXAccess

    ''' <summary>
    ''' Opérations de maintenance d'une base Access : compactage/réparation via DAO
    ''' (Microsoft Access Database Engine) en liaison tardive (late binding), sans
    ''' dépendance d'assembly d'interopérabilité.
    ''' </summary>
    Friend Module AccessMaintenance

        ''' <summary>
        ''' Compacte la base <paramref name="sourcePath"/>. Si <paramref name="targetPath"/>
        ''' est vide ou égal à la source, le compactage est effectué "en place" (via un
        ''' fichier temporaire puis remplacement atomique).
        ''' </summary>
        Friend Sub Compact(sourcePath As String, password As String, targetPath As String)
            If String.IsNullOrWhiteSpace(sourcePath) Then
                Throw New AccessConnectionException("Chemin de base introuvable pour le compactage (DatabasePath non défini).")
            End If
            If Not File.Exists(sourcePath) Then
                Throw New FileNotFoundException("Base de données Access introuvable.", sourcePath)
            End If

            Dim fullSource = Path.GetFullPath(sourcePath)
            Dim compactInPlace = String.IsNullOrWhiteSpace(targetPath) OrElse
                                 String.Equals(Path.GetFullPath(targetPath), fullSource, StringComparison.OrdinalIgnoreCase)

            Dim destination As String
            If compactInPlace Then
                Dim dir = Path.GetDirectoryName(fullSource)
                Dim tempName = $"{Path.GetFileNameWithoutExtension(fullSource)}_compact_{Guid.NewGuid():N}{Path.GetExtension(fullSource)}"
                destination = Path.Combine(dir, tempName)
            Else
                destination = Path.GetFullPath(targetPath)
            End If

            If File.Exists(destination) Then
                File.Delete(destination)
            End If

            CompactWithDao(fullSource, destination, password)

            If compactInPlace Then
                Dim backup = fullSource & ".bak_" & Guid.NewGuid().ToString("N")
                File.Replace(destination, fullSource, backup)
                Try
                    File.Delete(backup)
                Catch
                    ' La sauvegarde reste sur disque si elle ne peut être supprimée : non bloquant.
                End Try
            End If
        End Sub

        Private Sub CompactWithDao(source As String, destination As String, password As String)
            Dim daoType As Type = Type.GetTypeFromProgID("DAO.DBEngine.120")
            If daoType Is Nothing Then
                daoType = Type.GetTypeFromProgID("DAO.DBEngine.36")
            End If
            If daoType Is Nothing Then
                Throw New AccessConnectionException(
                    "Le moteur DAO (DAO.DBEngine.120) est introuvable. Installez 'Microsoft Access Database Engine' " &
                    "dans l'architecture du processus pour activer le compactage.")
            End If

            Dim engine As Object = Nothing
            Try
                engine = Activator.CreateInstance(daoType)

                Dim args As Object()
                If String.IsNullOrEmpty(password) Then
                    args = New Object() {source, destination}
                Else
                    ' Signature DAO : CompactDatabase(SrcName, DstName, [DstLocale], [Options], [SrcLocale])
                    ' Le mot de passe se transmet via le paramètre locale sous la forme ";pwd=".
                    Dim pwd = ";pwd=" & password
                    args = New Object() {source, destination, pwd, Type.Missing, pwd}
                End If

                daoType.InvokeMember("CompactDatabase", BindingFlags.InvokeMethod, Nothing, engine, args)
            Catch ex As TargetInvocationException When ex.InnerException IsNot Nothing
                Throw New AccessConnectionException(
                    "Échec du compactage de la base Access : " & ex.InnerException.Message, ex.InnerException)
            Finally
                If engine IsNot Nothing AndAlso Marshal.IsComObject(engine) Then
                    Marshal.FinalReleaseComObject(engine)
                End If
            End Try
        End Sub

    End Module

End Namespace
