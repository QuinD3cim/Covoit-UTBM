VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11430
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub B_Go_Click()
    Dim Address As String
    Dim Assdress2 As String
    Address = TB_address.Text
    Address2 = TB_address2.Text
    
    Call createVector(getCoordinates(Address), getCoordinates(Address2))
End Sub

Private Sub UserForm_Initialize()
   
End Sub

'getCoordinates|---Fonction permettant de récupérer les coordonées d'une addresse avec API mapquest

Private Function getCoordinates(Address) As String
    'Création des variables et constantes
    Dim Request As Object
    Dim Response() As String
    Dim API_Key As String
    Dim URL As String
    
    Dim x, y As String
    
    Dim flag1 As Boolean
    Dim flag2 As Boolean
    
    'Clef API Map Quest
    API_Key = "SHvZ3Q9XpGJufqwRn3isxllNhGJEsqiw"
    
    'URL en fonction de l'adresse
    URL = "http://www.mapquestapi.com/geocoding/v1/address?key=" & API_Key & "&outFormat=csv&location=" & Address & ",France" & "&maxResults=1&delimiter = %C2"
   
    'Création et envoi de la requête en GET
   
    Set Request = CreateObject("MSXML2.XMLHTTP")
    With Request
        .Open "GET", URL, False
        .send
    
    End With
    
    MsgBox Request.responseText
           
    'Gestion retour erreur de l'API
    
    If InStr(1, Request.responseText, "AppKey", vbTextCompare) Then
        MsgBox "Erreur de clef API contactez un admin"
        flag1 = True
    ElseIf InStr(1, Request.responseText, "400 Bad Request", vbTextCompare) Or Request.responseText = "" Then
        MsgBox "Erreur 400: Les caractères spéciaux et accents ne sont pas pris en compte"
        flag2 = True
    ElseIf (flag1 And flag2) = False Then
        'On récupère les informations sous forme de tableau
        
        Response = Split(Request.responseText, ",")
        'On supprime les guillemets en trop
        x = Replace(Response(UBound(Response) - 1), Chr(34), vbNullString)
        y = Replace(Response(UBound(Response)), Chr(34), vbNullString)

        'Format : (x;y)
        getCoordinates = "(" + x + ";" + y + ")"
        
        'Call createVector(getCoordinates, "(47.585829;6.865446)")
        
        Exit Function
        
        
    End If
    
    

'    For Each element In Response
'
'    MsgBox element
'
'    Next

End Function
'createVector|----Permet de créer un vecteur de déplacement entre 2 adresses
Private Function createVector(co_d, co_a) As String
    Dim co_departure, co_arrival() As String
    Dim x_departure, x_arrival, y_departure, y_arrival, x_vector, y_vector, abs_vector As Double
    
    
    co_d = Replace(co_d, Chr(40), vbNullString) 'supprimer ('
    co_d = Replace(co_d, Chr(41), vbNullString) 'supprimer )'
    co_d = Replace(co_d, Chr(46), Chr(44))    'convertir . -> ,'
    
    co_a = Replace(co_a, Chr(40), vbNullString) 'supprimer ('
    co_a = Replace(co_a, Chr(41), vbNullString) 'supprimer )'
    co_a = Replace(co_a, Chr(46), Chr(44))      'convertir . -> ,'
    
    'Coordonnées du point de Départ
    
    co_departure = Split(co_d, ";")
    
    x_departure = CDbl(co_departure(0))
    y_departure = CDbl(co_departure(1))
    
    'Coordonnées du point d'Arrivé
    
    co_arrival = Split(co_a, ";")
    
    x_arrival = CDbl(co_arrival(0))
    y_arrival = CDbl(co_arrival(1))
    
    'Calcul coordonnées et module du Vecteur
    
    x_vector = x_arrival - x_departure
    y_vector = y_arrival - y_departure
    
    abs_vector = 1.609344 * Sqr(x_vector ^ 2 + y_vector ^ 2) * 10 ^ 2
    
    'distance à vol d'oiseau
    createVector = Format(abs_vector, "####0.00")
    MsgBox Format(abs_vector, "####0.00") & " km | " & abs_vector & " km"
    Exit Function
    
    
End Function

