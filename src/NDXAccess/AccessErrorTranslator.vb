Imports System.Data.OleDb

Namespace NDXAccess

    ''' <summary>
    ''' Détecte les erreurs transitoires et traduit les <see cref="OleDbException"/> du
    ''' moteur ACE/Jet en messages clairs. La détection combine le code natif Jet/ACE
    ''' (souvent 0 avec le provider ACE) ET des motifs de message, plus fiables en pratique.
    ''' </summary>
    Friend Module AccessErrorTranslator

        ''' <summary>Codes d'erreur Jet/ACE de nature transitoire (verrou/contention).</summary>
        Private ReadOnly TransientCodes As New HashSet(Of Integer) From {
            3008, 3043, 3050, 3186, 3187, 3188, 3197, 3211, 3218, 3260, 3262
        }

        ''' <summary>Indique si l'exception OLE DB correspond à une erreur transitoire.</summary>
        Friend Function IsTransient(ex As OleDbException) As Boolean
            If ex Is Nothing Then Return False

            For Each err As OleDbError In ex.Errors
                If TransientCodes.Contains(err.NativeError) Then Return True
            Next

            Dim msg = ex.Message.ToLowerInvariant()
            Return msg.Contains("lock") OrElse msg.Contains("verrou") OrElse
                   msg.Contains("in use") OrElse msg.Contains("cours d'utilisation") OrElse
                   msg.Contains("déjà utilisé")
        End Function

        ''' <summary>Traduit une <see cref="OleDbException"/> en <see cref="AccessQueryException"/>.</summary>
        Friend Function Translate(ex As OleDbException) As AccessQueryException
            Dim code = GetPrimaryCode(ex)
            Dim message = Describe(code, ex)
            Return New AccessQueryException(message, code, IsTransient(ex), ex)
        End Function

        Private Function GetPrimaryCode(ex As OleDbException) As Integer
            If ex.Errors IsNot Nothing Then
                For Each err As OleDbError In ex.Errors
                    If err.NativeError <> 0 Then Return err.NativeError
                Next
            End If
            Return 0
        End Function

        Private Function Describe(code As Integer, ex As OleDbException) As String
            Dim hint = HintFromCode(code)
            If hint Is Nothing Then hint = HintFromMessage(ex.Message)

            If hint Is Nothing Then
                If code <> 0 Then
                    Return $"Erreur Access (code natif {code}) : {ex.Message}"
                End If
                Return $"Erreur Access : {ex.Message}"
            End If

            Dim codeText = If(code <> 0, $" (code ACE {code})", String.Empty)
            Return $"{hint}{codeText}. Détail d'origine : {ex.Message}"
        End Function

        Private Function HintFromCode(code As Integer) As String
            Select Case code
                Case 3031 : Return "Mot de passe de base de données invalide."
                Case 3033 : Return "Permissions insuffisantes sur la base ou l'objet."
                Case 3022 : Return "Violation de contrainte d'unicité (doublon de clé/index)."
                Case 3024 : Return "Fichier de base de données introuvable."
                Case 3044 : Return "Chemin de base de données invalide."
                Case 3049 : Return "Base endommagée, ou taille maximale (2 Go) atteinte."
                Case 3050 : Return "Impossible de verrouiller le fichier : verrou .laccdb détenu par un autre processus."
                Case 3051 : Return "Le fichier ne peut pas être ouvert : ouvert exclusivement ailleurs, ou droits insuffisants."
                Case 3197 : Return "Données modifiées par un autre utilisateur depuis leur lecture."
                Case 3201 : Return "Violation d'intégrité référentielle (enregistrement lié manquant)."
                Case 3211 : Return "Table verrouillée par un autre utilisateur."
                Case 3218, 3260, 3262 : Return "Enregistrement/table actuellement verrouillé par un autre utilisateur."
                Case 3343 : Return "Format de base de données non reconnu (fichier corrompu, ou version/provider incompatible)."
                Case Else : Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' Détection par motifs de message (le provider ACE renvoie fréquemment un code
        ''' natif 0 : le texte du message est alors la seule source fiable).
        ''' </summary>
        Private Function HintFromMessage(message As String) As String
            If String.IsNullOrEmpty(message) Then Return Nothing
            Dim m = message.ToLowerInvariant()

            If m.Contains("duplicate") OrElse m.Contains("doublon") OrElse m.Contains("en double") Then
                Return "Violation de contrainte d'unicité (doublon de clé/index)."
            ElseIf m.Contains("password") OrElse m.Contains("mot de passe") Then
                Return "Mot de passe de base de données invalide ou requis."
            ElseIf m.Contains("too few parameters") OrElse m.Contains("trop peu de paramètres") Then
                Return "Trop peu de paramètres : un nom de colonne/table est probablement mal orthographié (Access l'interprète comme un paramètre)."
            ElseIf m.Contains("lock") OrElse m.Contains("verrou") OrElse m.Contains("in use") OrElse m.Contains("cours d'utilisation") Then
                Return "Fichier ou enregistrement verrouillé par un autre utilisateur."
            ElseIf m.Contains("not a valid path") OrElse m.Contains("chemin") Then
                Return "Chemin de base de données invalide."
            ElseIf m.Contains("could not find") OrElse m.Contains("introuvable") OrElse m.Contains("cannot open") OrElse m.Contains("n'a pas pu trouver") Then
                Return "Fichier de base de données introuvable ou inaccessible."
            ElseIf m.Contains("unrecognized database format") OrElse m.Contains("format de base") OrElse m.Contains("non reconnu") Then
                Return "Format de base de données non reconnu (corruption, ou version/provider incompatible)."
            ElseIf m.Contains("syntax") OrElse m.Contains("syntaxe") Then
                Return "Erreur de syntaxe SQL."
            ElseIf m.Contains("relationship") OrElse m.Contains("integrity") OrElse m.Contains("intégrité") Then
                Return "Violation d'intégrité référentielle."
            End If

            Return Nothing
        End Function

    End Module

End Namespace
