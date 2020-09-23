Attribute VB_Name = "mdlRecherche"
Option Explicit

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public TrayIcon As NOTIFYICONDATA
Public affiche
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const FICHIER_TROUVE = " file(s) found"
Public Const TAILLEFICHIEROCTET = " koctets"
Public Const POURCENTAGE = " %"
Public Const VIDE = 0
Public Const CHAINE_VIDE = ""
Public Const DECALAGE_X = 145
Public Const DECALAGE_Y = 500
Public Const DECALAGE_BORDURE_X = 80
Public Const DECALAGE_BORDURE_Y = 760
Public Const DECALAGE_PLAYLIST_X = 6650
Public Const DECALAGE_PLAYLIST_BORDURE_X = 200
Public Const LIMITE_PLANTAGE_CONTROLE = 200
Public Const DECALAGE_HEIGHT = 420
Public Const DECALAGE_WIDTH = 120
Public Const VOLUME_MAX = 0
Public Const VOLUME_MID = -1500
Public Const VOLUME_MIN = -3000
Public Const VOLUME_STEP = 50
Public index_fichier_trouve As Long
Public Const ICO_MP3 = "C:\Mes documents\vbasic\Recherche\audiofile.ico"
Public Const ICO_TXT = "C:\Mes documents\vbasic\Recherche\txtfile.ico"
Public Const ICO_UNKNOWN = "C:\Mes documents\vbasic\Recherche\unknownfile.ico"
Public Const ICO_BMP = "C:\Mes documents\vbasic\Recherche\imagefile.ico"
Public Const ICO_EXE = "C:\Mes documents\vbasic\Recherche\exefile.ico"
Public Const ICO_AVI = "C:\Mes documents\vbasic\Recherche\videofile.ico"
Public Const ICO_APPLI = "C:\Mes documents\vbasic\Recherche\LITENING.ico"
Public Const BUT_PLAY = "C:\Mes documents\vbasic\Recherche\buttonPlay.bmp"
Public Const BUT_PAUSE = "C:\Mes documents\vbasic\Recherche\buttonPause.bmp"
Public Const BUT_STOP = "C:\Mes documents\vbasic\Recherche\buttonStop.bmp"

Public indicateur_recherche_en_cours As Boolean
Public Const DEBUT_PLAYLIST = 0
Public index_PlayList As Long
Public affich_menu As Boolean
Public MenuSys As Menu

Public pnl1 As Panel
Public pnl2 As Panel
Public Const ForReading = 1, ForWriting = 2, ForAppending = 3

Public Function ResetIndexFicherTrouve()

    index_fichier_trouve = VIDE

End Function

Public Function IncrementeIndexFichierTrouve()

    index_fichier_trouve = index_fichier_trouve + 1

End Function

Public Function FillToutTypeFichier(Control1 As ComboBox, Control2 As ComboBox, Control3 As ComboBox)

    With Control1
        .Text = "*.mp3"
        .AddItem "*.mp3"
        .AddItem "*.wav"
        .AddItem "*.txt"
        .AddItem "*.bmp"
        .AddItem "*.exe"
        .AddItem "*.avi"
        .AddItem "*.jpg"
    End With
    
    With Control2
        .Text = "Au moins"
        .AddItem "Au moins"
        .AddItem "Egal à"
        .AddItem "Au plus"
    End With
    
    With Control3
        .Text = "10 ko"
        .AddItem "10 ko"
        .AddItem "50 ko"
        .AddItem "100 ko"
        .AddItem "500 ko"
        .AddItem "1 mo"
        .AddItem "5 mo"
        .AddItem "10 mo"
        .AddItem "50 mo"
        .AddItem "100 mo"
        .AddItem "500 mo"
    End With
    
End Function

Public Function FillListBox(ByVal fichier As String, Control As TextBox)

    Dim Fs, File, temp As String
    
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set File = Fs.OpenTextFile(fichier, ForReading, 0)
    
    Do While File.AtEndOfStream <> True
        Control.Text = Control.Text & File.ReadLine & vbCrLf
    Loop
    
    File.Close
    
End Function

Public Function GetTitleFromFile(FileName As String, SearchChar As String) As String

    Dim i As Integer, j As Integer, strTemp As String
       
    i = 1
    If FileName <> "" Then
        While InStr(i, FileName, SearchChar, 1) > 0
            j = InStr(i, FileName, SearchChar, 1)
            i = j + 1
        Wend
        strTemp = Right(FileName, Len(FileName) - i + 1) '+1 -> afin de supprimer searchcar
        GetTitleFromFile = Left(strTemp, Len(strTemp) - 4) '4 = len(".???")
    Else
        GetTitleFromFile = ""
    End If
    
End Function

Public Function GetCombien_Octet(Texte As String) As Long

    If Right(Texte, 2) = "ko" Then
                            
        GetCombien_Octet = Left(Texte, (Len(Texte) - 3)) * 1000
                        
    ElseIf Right(Texte, 2) = "mo" Then
                        
        GetCombien_Octet = Left(Texte, (Len(Texte) - 3)) * 1000000
                        
    Else
                        
        GetCombien_Octet = Texte
                            
    End If

End Function

Sub PrintAnywhere(Src As Object, Dest As Object)
   
   Dest.PaintPicture Src.Picture, 0, 0
   
   If Dest Is Printer Then
      Printer.EndDoc
   End If
   
End Sub

Public Function NextZic(Control As MediaPlayer, list As ListBox)

    index_PlayList = index_PlayList + 1

    If index_PlayList >= list.ListCount Then
        
        'repetition
        index_PlayList = DEBUT_PLAYLIST
        Control.FileName = list.list(DEBUT_PLAYLIST)
        list.Selected(DEBUT_PLAYLIST) = True
        'rep
        
    ElseIf index_PlayList < list.ListCount Then
        list.Selected(index_PlayList) = True
        Control.FileName = list.list(index_PlayList)
    End If

End Function


