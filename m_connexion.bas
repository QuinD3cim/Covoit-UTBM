Attribute VB_Name = "m_connexion"
Global FichierAccess As Database
Global TableAccess As DAO.Recordset
Global rs As Recordset

Global table, requete As String
Global nbcomptes As Integer

Global id As Integer

'Cette procédure permet d'établir la liaiaon avec le fichier Access
Sub connexion()

    Set FichierAccess = OpenDatabase(ThisWorkbook.Path & "\Test.accdb")
    Set TableAccess = FichierAccess.OpenRecordset(table, dbOpenTable)
    
    MsgBox ("connexion avec la base de donnée établie")

End Sub

'Cette procédure permet la récupération de tous les identifiants, les noms d'utilisateur et tous les mots de passe
Sub recupid()

    'Récupération du nombre de comptes dans la base de données
    
    requete = "SELECT COUNT(*) FROM " & table
    Set rs = FichierAccess.OpenRecordset(requete)
    
    Range("A1").CopyFromRecordset rs
    nbcomptes = Range("A1").Value
    Range("A1").Value = ""
    
    'Récupérer les données d'identificaiton
    
    requete = "SELECT * FROM " & table
    Set rs = FichierAccess.OpenRecordset(requete)
    Range("A1").CopyFromRecordset rs

End Sub
