VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11430
   OleObjectBlob   =   "getGeocoding.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    

End Sub

Private Sub B_Go_Click()

    Dim Request As Object
    Dim Response() As String
    Dim URL As String
    
    URL = "http://www.mapquestapi.com/geocoding/v1/address?key=SHvZ3Q9XpGJufqwRn3isxllNhGJEsqiw&outFormat=csv&location=Washington,DC&delimiter=%C2"
    
   
    
   
    Set Request = CreateObject("MSXML2.XMLHTTP")
    With Request
        .Open "GET", URL, False
        .send
    
    End With
    'MsgBox Request.responseText
    
    Response = Split(Request.responseText, ",")
    MsgBox "(" & Response(UBound(Response) - 1) & "," & Response(UBound(Response)) & ")"
'
'    For Each element In Response
'
'    MsgBox element
'
'    Next

End Sub


