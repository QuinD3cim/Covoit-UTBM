VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_inscription 
   Caption         =   "Inscription"
   ClientHeight    =   3096
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9420.001
   OleObjectBlob   =   "uf_inscription.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_inscription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim utilisateur, mdp, naissance, numero, nom, prenom, utilisateurtrouve As String
Dim vehicule As Boolean


Sub change()

    b_valider.Enabled = True
    If ob_oui.Value = True Then
        vehicule = True
    Else
        vehicule = False
    End If
    
End Sub


Private Sub b_annuler_Click()

    If MsgBox("Voulez-vous vraiment annuler votre inscription ?", vbYesNo) = vbYes Then
        Dim ctrl As Control
        For Each ctrl In Me.Controls
           If TypeOf ctrl Is MSForms.TextBox Then
               ctrl.Text = ""
            ElseIf TypeOf ctrl Is MSForms.OptionButton Then
                ctrl.Value = False
           End If
        Next
        uf_inscription.Hide
        uf_connexion.Show
    End If
    
End Sub

Private Sub b_valider_Click()

    Dim erreur As Boolean
    Dim ctrl, ctrlerr As Control
    For Each ctrl In Me.Controls
        erreur = False
        If TypeOf ctrl Is MSForms.TextBox Then
            If Len(ctrl.Text) = 0 Then
                erreur = True
                Set ctrlerr = ctrl
                Exit For
            End If
        End If
    Next
    
    If erreur = True Or Len(tb_date.Text) <> 10 Or Len(tb_numero) <> 14 Then
        MsgBox "Vous n'avez pas rempli toutes les zones"
        ctrlerr.SetFocus
        Set ctrlerr = Nothing
    
    Else
    
        table = "connexion"
                        
        utilisateur = tb_utilisateur.Text
        mdp = tb_mdp.Text
        nom = tb_nom.Text
        prenom = tb_prenom.Text
        
        Call connexion
        Call recupid
            
        On Error GoTo nouveau
            utilisateurtrouve = Application.WorksheetFunction.VLookup(utilisateur, Range("B1:C" & nbcomptes), 2, False)
        
        If utilisateurtrouve = utilisateur Then
        
            MsgBox "Cet utilisateur existe déjà"
            End
        
        End If
        
nouveau:
        'ajout des identifiants dans le tableau connexion
        table = "connexion"
        Call connexion
        TableAccess.AddNew
        TableAccess.Fields("Utilisateur") = utilisateur
        TableAccess.Fields("Mot de passe") = mdp
        TableAccess.Update
        
        'ajout des informations personnelles dans le tableau profils
        table = "profils"
        Call connexion
        TableAccess.AddNew
        TableAccess.Fields("Prénom") = prenom
        TableAccess.Fields("Nom") = nom
        TableAccess.Fields("Date de naissance") = naissance
        TableAccess.Fields("Numéro de téléphone") = numero
        TableAccess.Fields("Véhicule") = vehicule
        TableAccess.Update
        
    End If
        
End Sub

Private Sub ob_non_Click()

    Call change
    
End Sub

Private Sub ob_oui_Change()
    
    Call change
    
End Sub

Private Sub tb_date_AfterUpdate()

    If IsDate(tb_date.Text) = False Then
        MsgBox "Date invalide", vbCritical, "Attention !"
        tb_date.Text = ""
    
    Else
        naissance = tb_date.Text
        
    End If

End Sub

Private Sub tb_date_Change()

    If IsNumeric(Right(tb_date.Text, 1)) = False And Len(tb_date.Text) <> 0 And Len(tb_date.Text) <> 3 And Len(tb_date.Text) <> 6 Then
        MsgBox "Entrez seulement des nombres"
        tb_date.Text = Left(tb_date.Text, Len(tb_date.Text) - 1)
        
    ElseIf Len(tb_date.Text) = 2 Or Len(tb_date.Text) = 5 Then
        tb_date.Text = tb_date.Text + "/"
    
    End If
    
    If Right(tb_date.Text, 1) = "/" And Len(naissance) = Len(tb_date.Text) + 1 Then
        tb_date.Text = Left(tb_date.Text, Len(tb_date.Text) - 1)
    End If
    
    naissance = tb_date.Text
    
End Sub

Private Sub tb_mdpverifie_Afterupdate()

    If tb_mdpverifie.Text <> tb_mdp Then
        MsgBox "Les mots de passe ne correspondent pas"
        tb_mdpverifie.Text = ""
    End If

End Sub

Private Sub tb_nom_Change()

    If Not Right(tb_nom.Text, 1) Like "[A-Z a-z - é è É È ê î ô]" Then
        If Len(tb_nom.Text) = 0 Then
            tb_nom.Text = ""
        Else
            tb_nom.Text = Left(tb_nom.Text, Len(tb_nom.Text) - 1)
        End If
    End If

End Sub

Private Sub tb_numero_Change()

    If IsNumeric(Right(tb_numero.Text, 1)) = False And Len(tb_numero.Text) <> 0 And (Len(tb_numero.Text) Mod 3) <> 0 Then
    
        MsgBox "Entrez seulement des nombres"
        
        tb_numero.Text = Left(tb_numero.Text, Len(tb_numero.Text) - 1)
        
    ElseIf (Len(tb_numero.Text) Mod 3) = 2 And Len(tb_numero.Text) <> 14 Then
        tb_numero.Text = tb_numero.Text + "."
    
    End If
    
    If Right(tb_numero.Text, 1) = "." And Len(numero) = Len(tb_numero.Text) + 1 Then
        tb_numero.Text = Left(tb_numero.Text, Len(tb_numero.Text) - 1)
    End If
    
    numero = tb_numero.Text

End Sub

Private Sub tb_prenom_Change()

    If Not Right(tb_prenom.Text, 1) Like "[A-Z a-z -]" Then
        If Len(tb_prenom.Text) = 0 Then
            tb_prenom.Text = ""
        Else
            tb_prenom.Text = Left(tb_prenom.Text, Len(tb_prenom.Text) - 1)
        End If
    End If
    
End Sub
