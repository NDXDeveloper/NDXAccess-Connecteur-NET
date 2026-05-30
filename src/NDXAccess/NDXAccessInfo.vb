Imports System.Reflection

Namespace NDXAccess

    ''' <summary>
    ''' Informations sur la bibliothèque NDXAccess (nom, version), lues depuis les
    ''' métadonnées de l'assembly. À ne pas confondre avec <see cref="DatabaseInfo.EngineVersion"/>
    ''' qui est la version du moteur Access (ACE/Jet).
    ''' </summary>
    Public Module NDXAccessInfo

        Private ReadOnly _assembly As Assembly = GetType(NDXAccessInfo).Assembly

        ''' <summary>Version de l'assembly (ex. "1.1.0.0").</summary>
        Public ReadOnly Property Version As String
            Get
                Return _assembly.GetName().Version?.ToString()
            End Get
        End Property

        ''' <summary>
        ''' Version informationnelle (ex. "1.1.0", éventuellement suffixée du hash de commit
        ''' en build CI). C'est la valeur recommandée pour l'affichage.
        ''' </summary>
        Public ReadOnly Property InformationalVersion As String
            Get
                Dim attr = _assembly.GetCustomAttribute(Of AssemblyInformationalVersionAttribute)()
                Return If(attr?.InformationalVersion, Version)
            End Get
        End Property

        ''' <summary>Nom du produit (ex. "NDXAccess").</summary>
        Public ReadOnly Property ProductName As String
            Get
                Dim attr = _assembly.GetCustomAttribute(Of AssemblyProductAttribute)()
                Return If(attr?.Product, _assembly.GetName().Name)
            End Get
        End Property

    End Module

End Namespace
