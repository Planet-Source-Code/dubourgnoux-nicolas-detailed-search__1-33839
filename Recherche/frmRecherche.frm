VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRecherche 
   Caption         =   "File search"
   ClientHeight    =   12960
   ClientLeft      =   6690
   ClientTop       =   4020
   ClientWidth     =   16890
   ForeColor       =   &H00404040&
   Icon            =   "frmRecherche.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12960
   ScaleWidth      =   16890
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   5040
      Top             =   9480
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   10800
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   31
      Top             =   1585
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   4680
      Top             =   8400
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5760
      Top             =   8520
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   16890
      _ExtentX        =   29792
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":158C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":19DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":2C60
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":326A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":3874
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":3E7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":4488
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":4A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":4DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":602E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":6480
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecherche.frx":68D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   12585
      Width           =   16890
      _ExtentX        =   29792
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.DirListBox dlbRecherche 
      BackColor       =   &H8000000D&
      ForeColor       =   &H0000FFFF&
      Height          =   1890
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   3135
   End
   Begin VB.ListBox lstHideRecherche 
      Height          =   645
      Left            =   6600
      TabIndex        =   9
      Top             =   8520
      Width           =   3015
   End
   Begin VB.FileListBox fleHideRecherche 
      Height          =   870
      Left            =   9720
      TabIndex        =   8
      Top             =   8280
      Width           =   2535
   End
   Begin VB.DriveListBox drvRecherche 
      BackColor       =   &H8000000D&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdRecherche 
      Caption         =   "&SEARCH"
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
   Begin VB.CheckBox chkVerifTaille 
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   1440
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Length"
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   6360
      TabIndex        =   15
      Top             =   1200
      Width           =   3135
      Begin VB.ComboBox cboVerif2Taille 
         BackColor       =   &H8000000D&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cboVerif1Taille 
         BackColor       =   &H8000000D&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   480
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ComboBox cboTypeFichier 
      BackColor       =   &H8000000D&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      ItemData        =   "frmRecherche.frx":7B54
      Left            =   3360
      List            =   "frmRecherche.frx":7B56
      TabIndex        =   2
      ToolTipText     =   "Types de fichiers connus"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtSearchName 
      BackColor       =   &H8000000D&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Details"
      ForeColor       =   &H00000040&
      Height          =   975
      Left            =   3240
      TabIndex        =   17
      Top             =   1200
      Width           =   2295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   0
      TabIndex        =   18
      Top             =   2640
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "Search results"
      TabPicture(0)   =   "frmRecherche.frx":7B58
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lswFound"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Playlist"
      TabPicture(1)   =   "frmRecherche.frx":7B74
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstPlayList"
      Tab(1).Control(1)=   "ProgressBar1"
      Tab(1).ControlCount=   2
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   -74930
         TabIndex        =   20
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Min             =   1.e-4
      End
      Begin VB.ListBox lstPlayList 
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   3420
         Left            =   -74940
         TabIndex        =   19
         Top             =   720
         Width           =   11175
      End
      Begin MSComctlLib.ListView lswFound 
         Height          =   4155
         Left            =   60
         TabIndex        =   28
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7329
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   65535
         BackColor       =   -2147483635
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fichier"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emplacement"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Taille"
            Object.Width           =   3351
         EndProperty
      End
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   615
      Left            =   9840
      Picture         =   "frmRecherche.frx":7B90
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Lecture"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pau&se"
      Height          =   615
      Left            =   10550
      Picture         =   "frmRecherche.frx":818A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Pause"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   615
      Left            =   11168
      Picture         =   "frmRecherche.frx":8784
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Stop"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   615
      Left            =   12000
      Picture         =   "frmRecherche.frx":8D7E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdClearPlayList 
      Caption         =   "E&rase"
      Height          =   615
      Left            =   12840
      Picture         =   "frmRecherche.frx":9378
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Efface la liste de lecture"
      Top             =   600
      Width           =   615
   End
   Begin VB.CheckBox chkMuteSound 
      Caption         =   "Mute"
      Height          =   255
      Left            =   9960
      TabIndex        =   29
      Top             =   1670
      Width           =   735
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   9840
      TabIndex        =   26
      ToolTipText     =   "Volume"
      Top             =   1200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Min             =   -3000
      Max             =   0
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   375
      Left            =   9840
      TabIndex        =   27
      ToolTipText     =   "Position"
      Top             =   2040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      SelectRange     =   -1  'True
      TickStyle       =   3
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      Height          =   2175
      Left            =   9720
      TabIndex        =   30
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   8280
      TabIndex        =   16
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   2040
      Width           =   1815
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   1215
      Left            =   5760
      TabIndex        =   13
      Top             =   5040
      Width           =   4335
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblNbFicherTrouve 
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Menu MonMenu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "&Ouvrir"
      End
      Begin VB.Menu mnuModify 
         Caption         =   "&Modifier"
      End
      Begin VB.Menu mnuAddPlayList 
         Caption         =   "&Ajouter à la liste de lecture"
      End
      Begin VB.Menu mnuAddAll 
         Caption         =   "Ajouter &tout"
      End
      Begin VB.Menu mnuLecture 
         Caption         =   "&Contrôles de lecture"
         Begin VB.Menu mnuPlay 
            Caption         =   "&Play"
         End
         Begin VB.Menu mnuPause 
            Caption         =   "Pau&se"
         End
         Begin VB.Menu mnuStop 
            Caption         =   "&Stop"
         End
         Begin VB.Menu mnuNext 
            Caption         =   "&Next"
         End
         Begin VB.Menu mnuErase 
            Caption         =   "E&rase"
         End
         Begin VB.Menu mnuMute 
            Caption         =   "&Mute"
            Checked         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmRecherche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkMuteSound_Click()

    Select Case chkMuteSound.Value
        Case vbChecked
            MediaPlayer1.Mute = True
        Case vbUnchecked
            MediaPlayer1.Mute = False
        Case Else
            Debug.Print "error mon senior"
    End Select

End Sub

Private Sub chkVerifTaille_Click()
    
    Select Case chkVerifTaille.Value
        Case vbChecked
            cboVerif1Taille.Enabled = True
            cboVerif2Taille.Enabled = True
        Case vbUnchecked
            cboVerif1Taille.Enabled = False
            cboVerif2Taille.Enabled = False
        Case Else
            Debug.Print "pas d'autre choix"
        End Select
            
End Sub

Private Sub cmdClearPlayList_Click()

    lstPlayList.Clear
    index_PlayList = VIDE
    Timer1.Enabled = False
    MediaPlayer1.FileName = ""
    ProgressBar1.Value = VIDE
    TrayIcon.hIcon = Me.Icon
    affiche = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
    Timer3.Enabled = False

End Sub

Private Sub cmdNext_Click()
 
    NextZic MediaPlayer1, lstPlayList
    
End Sub

Private Sub cmdPause_Click()
    
    If MediaPlayer1.FileName <> CHAINE_VIDE Then
        MediaPlayer1.Pause
        pnl2.Picture = LoadPicture(BUT_PAUSE)
        cmdPause.Enabled = False
        cmdPlay.Enabled = True
        cmdStop.Enabled = False
        Picture1.Visible = False
        TrayIcon.hIcon = Me.Icon
        affiche = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
        Timer3.Enabled = False
    End If
    
End Sub

Private Sub cmdPlay_Click()

    If MediaPlayer1.FileName <> CHAINE_VIDE Then
        MediaPlayer1.Play
        pnl2.Picture = LoadPicture(BUT_PLAY)
        cmdPlay.Enabled = False
        cmdPause.Enabled = True
        cmdStop.Enabled = True
        Picture1.Visible = True
        Timer3.Enabled = True
    End If

End Sub


Private Sub cmdRecherche_Click()

    drvRecherche.Enabled = False
    dlbRecherche.Enabled = False
    lswFound.Enabled = True
    
    Label1.Caption = "Searching..."
    Me.MousePointer = 13
    
    Dim i As Integer
    Dim temp As Long
    
    lstHideRecherche.Clear
    lswFound.ListItems.Clear
    ResetIndexFicherTrouve
    lblNbFicherTrouve.Caption = index_fichier_trouve & FICHIER_TROUVE
    Timer2.Enabled = True
    
    lstHideRecherche.AddItem dlbRecherche.Path
    
    i = 0
    temp = GetCombien_Octet(cboVerif2Taille.Text)

    If txtSearchName.Text = "" Then
    
        Do Until i >= lstHideRecherche.ListCount
            DoEvents
            dlbRecherche.Path = lstHideRecherche.list(i)
                    
            For j = 0 To fleHideRecherche.ListCount - 1
            
                If Right(LCase(fleHideRecherche.list(j)), 3) = Right(LCase(cboTypeFichier.Text), 3) Then
              
                    If chkVerifTaille.Value = vbChecked Then
                    
                        If cboVerif1Taille.Text = "Au moins" Then
                        
                            If FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) >= temp Then
                                    lswFound.ListItems.Add , , fleHideRecherche.list(j)
                                    lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(1) = dlbRecherche.Path
                                    lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(2) = FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) / 1000 & TAILLEFICHIEROCTET
                            
                                    IncrementeIndexFichierTrouve
                                    lblNbFicherTrouve.Caption = index_fichier_trouve & FICHIER_TROUVE
                
                            End If
                       
                        ElseIf cboVerif1Taille = "Egal à" Then
                            
                            If FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) = temp Then
                                lswFound.ListItems.Add , , fleHideRecherche.list(j)
                                lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(1) = dlbRecherche.Path
                                lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(2) = FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) / 1000 & TAILLEFICHIEROCTET
                                
                                IncrementeIndexFichierTrouve
                                lblNbFicherTrouve.Caption = index_fichier_trouve & FICHIER_TROUVE
                        
                            End If
                        
                        ElseIf cboVerif1Taille = "Au plus" Then
                    
                            If FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) <= temp Then
                                lswFound.ListItems.Add , , fleHideRecherche.list(j)
                                lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(1) = dlbRecherche.Path
                                lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(2) = FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) / 1000 & TAILLEFICHIEROCTET
                
                                IncrementeIndexFichierTrouve
                                lblNbFicherTrouve.Caption = index_fichier_trouve & FICHIER_TROUVE
                
                            End If
                        End If
                        
                    Else
                        lswFound.ListItems.Add , , fleHideRecherche.list(j)
                        lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(1) = dlbRecherche.Path
                        lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(2) = FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) / 1000 & TAILLEFICHIEROCTET
                                       
                        IncrementeIndexFichierTrouve
                        lblNbFicherTrouve.Caption = index_fichier_trouve & FICHIER_TROUVE
                                                  
                    End If
                
                End If
            
            Next

        For k = 0 To dlbRecherche.ListCount - 1
            lstHideRecherche.AddItem dlbRecherche.list(k)
        Next
        i = i + 1
    Loop
    
            dlbRecherche.Path = drvRecherche.Drive
            Label1.Caption = "Terminated"
            Me.MousePointer = 0
            Timer2.Enabled = False
            Label2.Caption = ""
    
            If index_fichier_trouve <> 0 Then
        
                If cboTypeFichier.Text = "*.mp3" Or cboTypeFichier = "*.wav" Then
                    lstPlayList.Visible = True
                End If
                
                affich_menu = True
            Else
                affich_menu = False
            End If
    
            drvRecherche.Enabled = True
            dlbRecherche.Enabled = True
    
    Else
              
        Do Until i >= lstHideRecherche.ListCount
            DoEvents
            dlbRecherche.Path = lstHideRecherche.list(i)
                    
            For j = 0 To fleHideRecherche.ListCount - 1
            
                If InStr(1, fleHideRecherche.list(j), txtSearchName, 1) <> 0 Then
                
                    If Right(LCase(fleHideRecherche.list(j)), 3) = Right(LCase(cboTypeFichier.Text), 3) Then
              
                        If chkVerifTaille.Value = vbChecked Then
                    
                            If cboVerif1Taille.Text = "Au moins" Then
                        
                                If FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) >= temp Then
                                    lswFound.ListItems.Add , , fleHideRecherche.list(j)
                                    lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(1) = dlbRecherche.Path
                                    lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(2) = FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) / 1000 & TAILLEFICHIEROCTET
                            
                                    IncrementeIndexFichierTrouve
                                    lblNbFicherTrouve.Caption = index_fichier_trouve & FICHIER_TROUVE
                
                                End If
                       
                            ElseIf cboVerif1Taille = "Egal à" Then
                            
                                If FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) = temp Then
                                    lswFound.ListItems.Add , , fleHideRecherche.list(j)
                                    lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(1) = dlbRecherche.Path
                                    lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(2) = FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) / 1000 & TAILLEFICHIEROCTET
                                
                                    IncrementeIndexFichierTrouve
                                    lblNbFicherTrouve.Caption = index_fichier_trouve & FICHIER_TROUVE
                                
                                End If
                        
                            ElseIf cboVerif1Taille = "Au plus" Then
                    
                                If FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) <= temp Then
                                    lswFound.ListItems.Add , , fleHideRecherche.list(j)
                                    lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(1) = dlbRecherche.Path
                                    lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(2) = FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) / 1000 & TAILLEFICHIEROCTET
                
                                    IncrementeIndexFichierTrouve
                                    lblNbFicherTrouve.Caption = index_fichier_trouve & FICHIER_TROUVE
                
                                End If
                            End If
                        
                        Else
                            lswFound.ListItems.Add , , fleHideRecherche.list(j)
                            lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(1) = dlbRecherche.Path
                            lswFound.ListItems.Item(lswFound.ListItems.Count).SubItems(2) = FileLen(dlbRecherche.Path & "\" & fleHideRecherche.list(j)) / 1000 & TAILLEFICHIEROCTET
                                       
                            IncrementeIndexFichierTrouve
                            lblNbFicherTrouve.Caption = index_fichier_trouve & FICHIER_TROUVE
                                                  
                        End If
                
                    End If
                    
                End If
            
            Next

            For k = 0 To dlbRecherche.ListCount - 1
                lstHideRecherche.AddItem dlbRecherche.list(k)
            Next
            i = i + 1
        Loop
    
        dlbRecherche.Path = drvRecherche.Drive
        Label1.Caption = "Recherche terminée"
        Me.MousePointer = 0
        Timer2.Enabled = False
        Label2.Caption = ""
    
        If index_fichier_trouve <> 0 Then
        
            If cboTypeFichier.Text = "*.mp3" Or cboTypeFichier = "*.wav" Then
                lstPlayList.Visible = True
            End If
                
            affich_menu = True
        Else
            affich_menu = False
        End If
    
        drvRecherche.Enabled = True
        dlbRecherche.Enabled = True
            
    End If
        
End Sub

Private Sub cmdStop_Click()

    If MediaPlayer1.FileName <> CHAINE_VIDE Then
        MediaPlayer1.Stop
        MediaPlayer1.CurrentPosition = VIDE
        Slider1.Value = VOLUME_MID
        pnl2.Picture = LoadPicture(BUT_STOP)
        cmdStop.Enabled = False
        cmdPlay.Enabled = True
        cmdPause.Enabled = False
        Picture1.Visible = False
        TrayIcon.hIcon = Me.Icon
        affiche = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
        Timer3.Enabled = False
    End If

End Sub





Private Sub dlbRecherche_Change()

    fleHideRecherche.Path = dlbRecherche.Path

End Sub

Private Sub drvRecherche_Change()

    On Error Resume Next
    dlbRecherche.Path = drvRecherche.Drive

End Sub

Private Sub Form_Load()
    
    'init de tray icon
    With TrayIcon
        .cbSize = Len(TrayIcon)
        .hWnd = Me.hWnd
        .uID = 1&
        .uFlags = &H1 Or &H2 Or &H4
        .uCallbackMessage = &H200
        .hIcon = Me.Icon
    End With

    affiche = Shell_NotifyIcon(NIM_ADD, TrayIcon)
    
    dlbRecherche.Path = "C:\"
    drvRecherche.Drive = dlbRecherche.Path
   
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    
    cmdPause.Enabled = False
    cmdStop.Enabled = False
    cmdNext.Enabled = False
    
    Slider1.Value = VOLUME_MID
    
    MediaPlayer1.Volume = VOLUME_MID
    MediaPlayer1.Visible = False

    lblNbFicherTrouve.Caption = index_fichier_trouve & FICHIER_TROUVE
    
    Set pnl1 = StatusBar1.Panels.Add()
    Set pnl2 = StatusBar1.Panels.Add()
    
    pnl1.Width = 1
    pnl2.Width = 1
    
    chkVerifTaille.Value = vbUnchecked
    cboVerif1Taille.Enabled = False
    cboVerif2Taille.Enabled = False
    
    fleHideRecherche.Visible = False
    lstHideRecherche.Visible = False
    lswFound.Enabled = False
    
    ProgressBar1.Visible = True
    ProgressBar1.Min = VIDE
        
    index_PlayList = VIDE
    indicateur_recherche_en_cours = False
    affich_menu = False
    chkMuteSound.Value = vbUnchecked
    mnuMute.Checked = False
    
    Picture1.Picture = LoadPicture(BUT_PLAY)
    Picture1.Visible = False
    
    mnuModify.Enabled = False 'temporaire
    
    ResetIndexFicherTrouve
    FillToutTypeFichier cboTypeFichier, cboVerif1Taille, cboVerif2Taille
            
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'If Button = vbRightButton Then Me.PopupMenu MenuSys

End Sub

Private Sub Form_Resize()

    If Me.WindowState <> vbMinimized Then
        With SSTab1
            .Width = Me.Width - DECALAGE_X
            If Me.Height - (StatusBar1.Height + DECALAGE_Y + .Top) > LIMITE_PLANTAGE_CONTROLE Then
                .Height = Me.Height - (StatusBar1.Height + DECALAGE_Y + .Top)
                lswFound.Width = .Width - DECALAGE_WIDTH
                    
                If .Height - DECALAGE_HEIGHT > 0 Then
                    If .Height - DECALAGE_HEIGHT - ProgressBar1.Height > 0 Then
                        lswFound.Height = .Height - DECALAGE_HEIGHT
                        lstPlayList.Width = .Width - DECALAGE_WIDTH
                        lstPlayList.Height = .Height - DECALAGE_HEIGHT - ProgressBar1.Height
                        ProgressBar1.Width = .Width - DECALAGE_WIDTH + DECALAGE_BORDURE_X
                    End If
                End If
            End If
            
        End With
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    affiche = Shell_NotifyIcon(NIM_DELETE, TrayIcon)
    Unload frmAffciheTexte

End Sub



Private Sub lstPlayList_DblClick()

    index_PlayList = lstPlayList.ListIndex
    MediaPlayer1.FileName = lstPlayList.list(index_PlayList)
    cmdPlay.Enabled = False
    cmdPause.Enabled = True
    cmdStop.Enabled = True
        
End Sub

Private Sub lstPlayList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    lstPlayList.ToolTipText = "Lecture en cours : " & index_PlayList + 1 & " @ " & GetTitleFromFile(lstPlayList.list(index_PlayList), "\")
    
End Sub

Private Sub lstPlayList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then Me.PopupMenu mnuLecture

End Sub

Private Sub lswFound_Click()
    
    If index_fichier_trouve <> VIDE Then
        pnl1.AutoSize = sbrContents
        pnl1.Text = lswFound.SelectedItem.Text
                  
        Select Case Right(LCase(lswFound.SelectedItem), 3)
            Case "mp3"
                pnl1.Picture = LoadPicture(ICO_MP3)
                mnuAddPlayList.Enabled = True
                mnuAddAll.Enabled = True
            Case "wav"
                pnl1.Picture = LoadPicture(ICO_MP3)
                mnuAddPlayList.Enabled = True
                mnuAddAll.Enabled = True
            Case "txt"
                pnl1.Picture = LoadPicture(ICO_TXT)
                mnuAddPlayList.Enabled = False
                mnuAddAll.Enabled = False
            Case "bmp"
                pnl1.Picture = LoadPicture(ICO_BMP)
                mnuAddPlayList.Enabled = False
                mnuAddAll.Enabled = False
            Case "jpg"
                pnl1.Picture = LoadPicture(ICO_BMP)
                mnuAddPlayList.Enabled = False
                mnuAddAll.Enabled = False
            Case "exe"
                pnl1.Picture = LoadPicture(ICO_EXE)
                mnuAddPlayList.Enabled = False
                mnuAddAll.Enabled = False
            Case "avi"
                pnl1.Picture = LoadPicture(ICO_AVI)
                mnuAddPlayList.Enabled = False
                mnuAddAll.Enabled = False
            Case Else
                pnl1.Picture = LoadPicture(ICO_UNKNOWN)
                mnuAddPlayList.Enabled = False
                mnuAddAll.Enabled = False
                mnuOpen.Enabled = False
            End Select
    End If

End Sub

Private Sub lswFound_DblClick()
  
    If index_fichier_trouve <> VIDE Then
        pnl1.AutoSize = sbrContents
        pnl1.Text = lswFound.SelectedItem.Text
                  
        Select Case Right(LCase(lswFound.SelectedItem), 3)
            Case "mp3"
                index_PlayList = DEBUT_PLAYLIST
                lstPlayList.Clear
                lstPlayList.list(DEBUT_PLAYLIST) = lswFound.SelectedItem.SubItems(1) & "\" & lswFound.SelectedItem
                MediaPlayer1.FileName = lstPlayList.list(DEBUT_PLAYLIST)
                pnl1.Picture = LoadPicture(ICO_MP3)
                pnl2.Picture = LoadPicture(BUT_PLAY)
                cmdPlay.Enabled = False
                cmdPause.Enabled = True
                cmdStop.Enabled = True
                cmdNext.Enabled = True
                lstPlayList.Selected(0) = True
                Timer3.Enabled = True
            Case "wav"
                index_PlayList = DEBUT_PLAYLIST
                lstPlayList.Clear
                lstPlayList.list(DEBUT_PLAYLIST) = lswFound.SelectedItem.SubItems(1) & "\" & lswFound.SelectedItem
                MediaPlayer1.FileName = lstPlayList.list(DEBUT_PLAYLIST)
                pnl1.Picture = LoadPicture(ICO_MP3)
                pnl2.Picture = LoadPicture(BUT_PLAY)
                cmdPlay.Enabled = False
                cmdPause.Enabled = True
                cmdStop.Enabled = True
                cmdNext.Enabled = True
                lstPlayList.Selected(0) = True
                Timer3.Enabled = True
            Case "txt"
                pnl1.Picture = LoadPicture(ICO_TXT)
                frmAffciheTexte.Show vbModal
            Case "bmp"
                pnl1.Picture = LoadPicture(ICO_BMP)
                frmAfficheImage.Show vbModal
            Case "jpg"
                pnl1.Picture = LoadPicture(ICO_BMP)
                frmAfficheImage.Show vbModal
            Case "exe"
                pnl1.Picture = LoadPicture(ICO_EXE)
                Shell lswFound.SelectedItem.SubItems(1) & "\" & lswFound.SelectedItem, vbNormalFocus
            Case "avi"
                pnl1.Picture = LoadPicture(ICO_AVI)
                frmLireVideo.Show vbModal
            Case Else
                pnl1.Picture = LoadPicture(ICO_UNKNOWN)
            End Select
    End If
    
End Sub

Private Sub lswFound_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Button
        Case vbLeftButton
            '????
        Case vbRightButton
            If affich_menu = True Then
                If Right(LCase(lswFound.SelectedItem), 3) = "mp3" Then
                    mnuAddPlayList.Enabled = True
                    mnuAddAll.Enabled = True
                    mnuOpen.Enabled = True
                ElseIf Right(LCase(lswFound.SelectedItem), 3) = "wav" Then
                    mnuAddPlayList.Enabled = True
                    mnuAddAll.Enabled = True
                Else
                    mnuAddPlayList.Enabled = False
                    mnuAddAll.Enabled = False
                End If
                Me.PopupMenu MonMenu
            End If
        End Select

End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)

    cmdPause.Enabled = False
    cmdStop.Enabled = False
    cmdPlay.Enabled = True
        
    Timer1.Enabled = False
    pnl2.Text = "Terminé"
    
    NextZic MediaPlayer1, lstPlayList
    
    cmdPause.Enabled = True
    cmdStop.Enabled = True
    cmdNext.Enabled = True
    cmdPlay.Enabled = False
        

End Sub

Private Sub MediaPlayer1_NewStream()

    Slider2.Max = MediaPlayer1.Duration
    Slider2.Value = 0

End Sub

Private Sub MediaPlayer1_PlayStateChange(ByVal OldState As Long, ByVal newState As Long)
    
    Timer1.Enabled = True

End Sub



Private Sub mnuAddAll_Click()

    Dim i As Long
    
    ProgressBar1.Visible = True
    
    For i = 1 To lswFound.ListItems.Count
        lstPlayList.AddItem lswFound.ListItems.Item(i).ListSubItems(1) & "\" & lswFound.ListItems.Item(i).Text
    Next i
    
    MediaPlayer1.FileName = lstPlayList.list(0)
    ProgressBar1.Max = lstPlayList.ListCount
    
    cmdNext.Enabled = True
    cmdPlay.Enabled = False
    cmdStop.Enabled = True
    cmdPause.Enabled = True
    
    Timer3.Enabled = True

End Sub

Private Sub mnuAddPlayList_Click()

    If lswFound.Enabled = True Then
            If lstPlayList.ListCount = 0 Then
                MediaPlayer1.FileName = lswFound.SelectedItem.SubItems(1) & "\" & lswFound.SelectedItem
                cmdNext.Enabled = True
                cmdPlay.Enabled = False
                cmdStop.Enabled = True
                cmdPause.Enabled = True
            End If
            
            lstPlayList.AddItem lswFound.SelectedItem.SubItems(1) & "\" & lswFound.SelectedItem
            ProgressBar1.Visible = True
            ProgressBar1.Max = lstPlayList.ListCount
            
            Timer3.Enabled = True
    End If

End Sub

Private Sub mnuMute_Click()

    If mnuMute.Checked = False Then
        chkMuteSound.Value = vbChecked
        MediaPlayer1.Mute = True
        mnuMute.Checked = True
    Else
        chkMuteSound.Value = vbUnchecked
        MediaPlayer1.Mute = False
        mnuMute.Checked = False
    End If

End Sub

Private Sub mnuOpen_Click()

    lswFound_DblClick
    ProgressBar1.Max = lstPlayList.ListCount

End Sub

Private Sub mnuPause_Click()

    'mnuPause_Click

End Sub

Private Sub mnuPlay_Click()

    'cmdPlay_Click
        
End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MediaPlayer1.Volume = Slider1.Value

End Sub

Private Sub Slider2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Timer1.Enabled = False

End Sub

Private Sub Slider2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MediaPlayer1.CurrentPosition = Slider2.Value
    Timer1.Enabled = True

End Sub



Private Sub Timer1_Timer()

    If MediaPlayer1.Duration > 0 Then
        pnl2.AutoSize = sbrContents
        pnl2.Text = "Progression de [" & _
                    GetTitleFromFile(MediaPlayer1.FileName, "\") & "]" & " : " & _
                    Format(((MediaPlayer1.CurrentPosition * 100) / MediaPlayer1.Duration), "##0") & POURCENTAGE
        
        lstPlayList.Selected(index_PlayList) = True
        
        Slider2.Value = MediaPlayer1.CurrentPosition
        Slider2.ToolTipText = "Progression : " & _
                            Format(((MediaPlayer1.CurrentPosition * 100) / MediaPlayer1.Duration), "##0") & POURCENTAGE
            
        ProgressBar1.Value = index_PlayList
        ProgressBar1.ToolTipText = "Progression de la liste de lecture : " & index_PlayList + 1 & " sur " & lstPlayList.ListCount
                
    End If
        
End Sub

Private Sub Timer2_Timer()

    If indicateur_recherche_en_cours = False Then
        Label2.Caption = "\"
        indicateur_recherche_en_cours = True
    Else
        Label2.Caption = "/"
        indicateur_recherche_en_cours = False
    End If

End Sub

Private Sub Timer3_Timer()

    If Picture1.Visible = True Then
        Picture1.Visible = False
    Else
        Picture1.Visible = True
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Index = 2 Then
        cmdRecherche_Click
    ElseIf Button.Index = 4 Then
        frmAbout.Show vbModal
    End If

End Sub

