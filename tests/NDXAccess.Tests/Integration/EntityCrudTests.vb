Imports FluentAssertions
Imports NDXAccess
Imports NDXAccess.Examples
Imports Xunit

''' <summary>
''' Tests d'intégration de l'exemple d'entité "Active Record" (Etape) — exécutés
''' réellement contre une base .accdb. Crée/supprime son propre schéma (tb_Etapes,
''' tb_Modules, tb_Etape_Modules) requis par l'entité.
''' </summary>
<Collection("Access")>
<Trait("Category", "Integration")>
Public Class EntityCrudTests

    Private ReadOnly _fixture As AccessDatabaseFixture

    Public Sub New(fixture As AccessDatabaseFixture)
        _fixture = fixture
    End Sub

    Private Shared Sub EnsureSchema(c As IAccessConnection)
        DropSchema(c)
        c.ExecuteNonQuery("CREATE TABLE tb_Etapes (id AUTOINCREMENT PRIMARY KEY, Libelle TEXT(100), Numero LONG, Fonction LONG, fkChapitre LONG, DateCrea DATETIME, DateModif DATETIME, ParQui TEXT(100))")
        c.ExecuteNonQuery("CREATE TABLE tb_Modules (id AUTOINCREMENT PRIMARY KEY, Libelle TEXT(100))")
        c.ExecuteNonQuery("CREATE TABLE tb_Etape_Modules (id AUTOINCREMENT PRIMARY KEY, fkEtape LONG, fkModl LONG)")
    End Sub

    Private Shared Sub DropSchema(c As IAccessConnection)
        For Each t In {"tb_Etape_Modules", "tb_Etapes", "tb_Modules"}
            Try : c.ExecuteNonQuery($"DROP TABLE {t}") : Catch : End Try
        Next
    End Sub

    <SkippableFact>
    Public Sub Save_Insert_ShouldAssignNewId()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Using c = _fixture.CreateConnection()
            EnsureSchema(c)
            Try
                Dim e As New Etape(c) With {
                    .Libelle = "Préparation",
                    .Numero = 1,
                    .Fonction = Etape.EnumFonction.Auto,
                    .FkChapitre = 10
                }
                e.Save().Should().BeTrue()
                e.Id.Should().BeGreaterThan(0)
            Finally
                DropSchema(c)
            End Try
        End Using
    End Sub

    <SkippableFact>
    Public Sub LoadFromID_ShouldReloadSavedEntity()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Using c = _fixture.CreateConnection()
            EnsureSchema(c)
            Try
                Dim e As New Etape(c) With {.Libelle = "Etape A", .Numero = 5, .FkChapitre = 2, .Fonction = Etape.EnumFonction.Bloquante}
                e.Save().Should().BeTrue()
                Dim id = e.Id

                Dim loaded As New Etape(c)
                loaded.LoadFromID(id).Should().BeTrue()
                loaded.Libelle.Should().Be("Etape A")
                loaded.Numero.Should().Be(5)
                loaded.FkChapitre.Should().Be(2)
                loaded.Fonction.Should().Be(Etape.EnumFonction.Bloquante)
            Finally
                DropSchema(c)
            End Try
        End Using
    End Sub

    <SkippableFact>
    Public Sub Save_Update_ShouldModifyExistingRow()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Using c = _fixture.CreateConnection()
            EnsureSchema(c)
            Try
                Dim e As New Etape(c) With {.Libelle = "Avant", .Numero = 1, .FkChapitre = 1}
                e.Save().Should().BeTrue()
                Dim id = e.Id

                e.Libelle = "Après"
                e.Save().Should().BeTrue()
                e.Id.Should().Be(id)   ' toujours la même ligne (UPDATE)

                Dim reloaded As New Etape(c)
                reloaded.LoadFromID(id).Should().BeTrue()
                reloaded.Libelle.Should().Be("Après")

                ' Une seule ligne au total.
                c.ExecuteScalar(Of Integer)("SELECT COUNT(*) FROM tb_Etapes").Should().Be(1)
            Finally
                DropSchema(c)
            End Try
        End Using
    End Sub

    <SkippableFact>
    Public Sub LoadFrom_ByNumeroAndChapitre_ShouldWork()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Using c = _fixture.CreateConnection()
            EnsureSchema(c)
            Try
                Dim e As New Etape(c) With {.Libelle = "Cible", .Numero = 7, .FkChapitre = 3}
                e.Save().Should().BeTrue()

                Dim found As New Etape(c)
                found.LoadFrom(7, 3).Should().BeTrue()
                found.Id.Should().Be(e.Id)
                found.Libelle.Should().Be("Cible")
            Finally
                DropSchema(c)
            End Try
        End Using
    End Sub

    <SkippableFact>
    Public Sub Remove_ShouldDeleteRow()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Using c = _fixture.CreateConnection()
            EnsureSchema(c)
            Try
                Dim e As New Etape(c) With {.Libelle = "À supprimer", .Numero = 1, .FkChapitre = 1}
                e.Save().Should().BeTrue()
                Dim id = e.Id

                e.Remove().Should().BeTrue()

                Dim gone As New Etape(c)
                gone.LoadFromID(id).Should().BeFalse()
                c.ExecuteScalar(Of Integer)("SELECT COUNT(*) FROM tb_Etapes").Should().Be(0)
            Finally
                DropSchema(c)
            End Try
        End Using
    End Sub

    <SkippableFact>
    Public Sub ListeModules_ShouldLazyLoadChildren()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Using c = _fixture.CreateConnection()
            EnsureSchema(c)
            Try
                ' Un module
                c.ExecuteNonQuery("INSERT INTO tb_Modules (Libelle) VALUES (?)", "Module 1")
                Dim modId = c.ExecuteScalar(Of Integer)("SELECT @@IDENTITY")

                ' Une étape
                Dim e As New Etape(c) With {.Libelle = "Avec modules", .Numero = 1, .FkChapitre = 1}
                e.Save().Should().BeTrue()

                ' Lien étape <-> module
                c.ExecuteNonQuery("INSERT INTO tb_Etape_Modules (fkEtape, fkModl) VALUES (?, ?)", e.Id, modId)

                ' Chargement paresseux de la collection enfant
                e.ListeModules.Count.Should().Be(1)
                e.ListeModules(0).Id.Should().Be(modId)
                e.ListeModules(0).Libelle.Should().Be("Module 1")
            Finally
                DropSchema(c)
            End Try
        End Using
    End Sub

    <SkippableFact>
    Public Sub LoadFromIDViaConnector_ShouldWork()
        Skip.IfNot(_fixture.ProviderAvailable, _fixture.SkipReason)
        Using c = _fixture.CreateConnection()
            EnsureSchema(c)
            Try
                Dim e As New Etape(c) With {.Libelle = "Via connecteur", .Numero = 9, .FkChapitre = 4}
                e.Save().Should().BeTrue()

                Dim loaded As New Etape(c)
                loaded.LoadFromIDViaConnector(e.Id).Should().BeTrue()
                loaded.Libelle.Should().Be("Via connecteur")
            Finally
                DropSchema(c)
            End Try
        End Using
    End Sub

End Class
