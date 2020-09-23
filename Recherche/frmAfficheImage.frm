VERSION 5.00
Begin VB.Form frmAfficheImage 
   AutoRedraw      =   -1  'True
   ClientHeight    =   9795
   ClientLeft      =   7710
   ClientTop       =   5685
   ClientWidth     =   13005
   Icon            =   "frmAfficheImage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   13005
   Begin VB.Menu mnuP 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuS1 
         Caption         =   "&Imprimer"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuS2 
         Caption         =   "&Quitter"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmAfficheImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    
    Dim fileImage As String
    
    fileImage = frmRecherche.lswFound.SelectedItem.SubItems(1) & "\" & frmRecherche.lswFound.SelectedItem
    
    Me.Caption = fileImage
       
    Me.Picture = LoadPicture(fileImage)
        
End Sub

Private Sub mnuS1_Click()

    PrintAnywhere frmAfficheImage, Printer

End Sub

Private Sub mnuS2_Click()

    Unload Me

End Sub
