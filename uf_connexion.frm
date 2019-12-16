VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_connexion 
   Caption         =   "Connexion Covoit'UTBM"
   ClientHeight    =   3144
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5928
   OleObjectBlob   =   "uf_connexion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_connexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim utilisateur, mdp, utilisateurtrouve As String
Dim numutilisateur As Integer

Private Sub b_connecter_Click()

    Call connexion
    Call recupid
    
    'Vérifie si le compte entré est dans la base de donnée
    
    utilisateur = tb_utilisateur.Text
    mdp = tb_mdp.Text
    
    On Error GoTo erreur
    utilisateurtrouve = Application.WorksheetFunction.VLookup(utilisateur, Range("B1:C" & nbcomptes), 2, False)
    
    If utilisateurtrouve = mdp Then
        
        For i = 1 To nbcomptes
            If Range("B" & i).Text = utilisateur Then
                id = Range("A" & i).Value
                Exit For
            End If
        Next
        
        MsgBox "Vous êtes connecté, bienvenue " & id
        
    Else:
        MsgBox "Mot de passe erroné"
        
    End If
    
    If False Then
erreur:
    MsgBox "Utilisateur inconnu"
    End If
    
    
End Sub


Private Sub b_fin_Click()

    If MsgBox("Vous êtes sûr de vouloir quitter ?", vbYesNo) = vbYes Then
        End
    End If
    
End Sub

Private Sub b_inscription_Click()

    uf_connexion.Hide
    uf_inscription.Show
    
End Sub

Private Sub UserForm_Initialize()

    Worksheets("info").Activate
    
    Cells.Select
    Selection.Font.ColorIndex = 1
    
    table = "connexion"
    tb_mdp.PasswordChar = "*"

End Sub

