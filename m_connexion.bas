Attribute VB_Name = "m_connexion"
Global FichierAccess As Database
Global TableAccess As DAO.Recordset
Global rs As Recordset

Global table, requete As String
Global nbcomptes As Integer

Global id As Integer

'Cette proc�dure permet d'�tablir la liaiaon avec le fichier Access
Sub connexion()

    Set FichierAccess = OpenDatabase(ThisWorkbook.Path & "\Test.accdb")
    Set TableAccess = FichierAccess.OpenRecordset(table, dbOpenTable)
    
    MsgBox ("connexion avec la base de donn�e �tablie")

End Sub

'Cette proc�dure permet la r�cup�ration de tous les identifiants, les noms d'utilisateur et tous les mots de passe
Sub recupid()

    'R�cup�ration du nombre de comptes dans la base de donn�es
    
    requete = "SELECT COUNT(*) FROM " & table
    Set rs = FichierAccess.OpenRecordset(requete)
    
    Range("A1").CopyFromRecordset rs
    nbcomptes = Range("A1").Value
    Range("A1").Value = ""
    
    'R�cup�rer les donn�es d'identificaiton
    
    requete = "SELECT * FROM " & table
    Set rs = FichierAccess.OpenRecordset(requete)
    Range("A1").CopyFromRecordset rs

End Sub
