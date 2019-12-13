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
Dim FichierAccess As Database
Dim TableAcces As DAO.Recordset
Dim rs As Recordset

Dim table As String
Dim requete As String
Dim nbcomptes As Integer

Sub connexion()

    Set FichierAccess = OpenDatabase("Covoitutbm.accdb")
    Set TableAccess = FichierAccess.OpenRecordset("connexion", dbOpenTable)
    
    MsgBox ("connexion établie")

End Sub

Private Sub b_connecter_Click()

    Call connexion
    
    'Récupération du nombre de comptes dans la base de données
    
    requete = "SELECT COUNT(*) FROM " & table
    Set rs = FichierAccess.OpenRecordset(requete)
    
    Range("A1").CopyFromRecordset rs
    nbcomptes = Range("A1").Value
    Range("A1").Value = ""
    
    'Récupérer les identifiants des comptes
    
    requete = "SELECT * FROM table_name"
    
    For i = 1 To nbcomptes
    
        requete = "SELECT " &  & " FROM " & table
        Range("A"&i).Value =
    
    
End Sub


Private Sub UserForm_Initialize()

    table = "connexion"
    tb_mdp.PasswordChar = "*"

End Sub

