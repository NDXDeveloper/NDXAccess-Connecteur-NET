Imports System.Text

Namespace NDXAccess

    ''' <summary>
    ''' Traduit un SQL à paramètres nommés (@nom) en SQL positionnel OLE DB (?), en
    ''' produisant le tableau de valeurs dans l'ordre d'apparition.
    ''' </summary>
    ''' <remarks>
    ''' <para>Gère les littéraux chaîne entre apostrophes (les @ y sont ignorés).</para>
    ''' <para>Un même paramètre utilisé plusieurs fois est ajouté autant de fois (OLE DB
    ''' étant positionnel). Les jetons inconnus (ex. <c>@@IDENTITY</c>) sont laissés tels quels.</para>
    ''' </remarks>
    Friend Module NamedParameterParser

        Friend Function Translate(sql As String, params As IDictionary(Of String, Object)) As (Sql As String, Values As Object())
            If String.IsNullOrEmpty(sql) OrElse params Is Nothing OrElse params.Count = 0 Then
                Return (sql, Array.Empty(Of Object)())
            End If

            Dim sb As New StringBuilder(sql.Length)
            Dim values As New List(Of Object)()
            Dim i = 0
            Dim inString = False

            While i < sql.Length
                Dim c = sql(i)

                If c = "'"c Then
                    inString = Not inString
                    sb.Append(c)
                    i += 1
                ElseIf (Not inString) AndAlso c = "@"c Then
                    ' Lire l'identifiant qui suit le @.
                    Dim j = i + 1
                    While j < sql.Length AndAlso (Char.IsLetterOrDigit(sql(j)) OrElse sql(j) = "_"c)
                        j += 1
                    End While

                    Dim withAt = sql.Substring(i, j - i)      ' inclut le @
                    Dim key = withAt.Substring(1)              ' sans le @
                    Dim value As Object = Nothing

                    If j > i + 1 AndAlso (params.TryGetValue(key, value) OrElse params.TryGetValue(withAt, value)) Then
                        sb.Append("?"c)
                        values.Add(value)
                        i = j
                    Else
                        ' Jeton non reconnu (ex. @@IDENTITY) : laissé tel quel.
                        sb.Append(c)
                        i += 1
                    End If
                Else
                    sb.Append(c)
                    i += 1
                End If
            End While

            Return (sb.ToString(), values.ToArray())
        End Function

    End Module

End Namespace
