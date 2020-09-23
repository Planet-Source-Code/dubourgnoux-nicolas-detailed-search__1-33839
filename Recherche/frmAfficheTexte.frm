VERSION 5.00
Begin VB.Form frmAffciheTexte 
   ClientHeight    =   5775
   ClientLeft      =   8595
   ClientTop       =   7275
   ClientWidth     =   9750
   Icon            =   "frmAfficheTexte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9750
   Begin VB.TextBox txtText 
      Height          =   1335
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   9735
   End
   Begin VB.Menu mnuP 
      Caption         =   "&Fichier"
      Index           =   0
      WindowList      =   -1  'True
      Begin VB.Menu mnuS1 
         Caption         =   "&Imprimer"
         Index           =   1
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuS2 
         Caption         =   "&Fermer"
         Index           =   2
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "frmAffciheTexte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    With txtText
        .Height = frmRecherche.Height
        .Width = frmRecherche.Width
    End With
    
    Me.Caption = frmRecherche.lswFound.SelectedItem.SubItems(1) & "\" & frmRecherche.lswFound.SelectedItem
    
    FillListBox Me.Caption, txtText

End Sub

Private Sub Form_Resize()
    
    With txtText
        .Height = Me.Height - DECALAGE_BORDURE_Y
        .Width = Me.Width - DECALAGE_BORDURE_X
    End With
    
End Sub


Private Sub mnuS1_Click(Index As Integer)

  Printer.Print txtText.Text
  Printer.NewPage
           
End Sub

Private Sub mnuS2_Click(Index As Integer)

    Unload Me

End Sub
