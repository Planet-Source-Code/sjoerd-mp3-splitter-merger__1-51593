VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmJoin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sjoerd MP3 Merger"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl m 
      Height          =   330
      Left            =   840
      TabIndex        =   16
      Top             =   4320
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.TextBox txtGenre 
      Height          =   285
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   15
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txtComment 
      Height          =   285
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   14
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   13
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtAlbum 
      Height          =   285
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   12
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   11
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   10
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "&Join"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox lstFiles 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   2640
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblGenre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Genre:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   480
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   705
   End
   Begin VB.Label lblYear 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label lblAlbum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Album:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label lblArtist 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Artist:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   390
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   345
   End
End
Attribute VB_Name = "frmJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hi,
'This code can merge multiple MP3 files to one single MP3 file
'Just run the code
'Click Add and select the files you wan't to merge in a single MP3
'Click Clear to clear the list
'After your list contains all the files you wan't edit the IDE 3 Tag
'And after that/ or click Join to create one single MP3
'-------------------------------------------------------------------
'Some MP3 players aren't able to play the newly created MP3 file
'This occurs when you merge MP3's containing the IDE 3 V.2 Tag
'This code only deletes the IDE 3 V.1 Tag, coz the V.2 Tag can be any size
'and can occur on almost any place in the MP3 file
'When you play this file, some players will think that a new MP3 started
'and the MP3 player will stop
'-------------------------------------------------------------------
'The 'WriteTag' function is from another coder (see clsMp3 for the Pscode link
'U can use this code for whatever you wan't, but please leave credits for me (Sjoerd) or
'when you use the class or the 'WriteTag' function leave credits for this coder
'Thanks and much fun!!!
'Greetings, Sjoerd
'Please vote for me and leave comments

Option Explicit

Private Sub cmdAdd_Click()

    cDialog.FileName = Empty
    cDialog.ShowOpen
    If LenB(cDialog.FileName) = 0 Then
        Exit Sub
    End If
    lstFiles.AddItem cDialog.FileName

End Sub

Private Sub cmdClear_Click()

    lstFiles.Clear

End Sub

Private Sub cmdJoin_Click()

 
  Dim FileName1
  Dim I            As Integer
  Dim IDE3         As New clsMp3
  Dim strData      As String
  Dim strAll       As String
  Dim intI         As Integer
    Me.Caption = "Selecting destination file..."
    cDialog.FileName = Empty
    cDialog.ShowSave
    FileName1 = cDialog.FileName
    If LenB(FileName1) = 0 Then
        Exit Sub
    End If
    For I = 0 To lstFiles.ListCount - 1
        DoEvents
        Me.Caption = "Reading IDE3 tag..."
        IDE3.ReadMP3 lstFiles.List(I)
        intI = FreeFile
        Open lstFiles.List(I) For Binary As #1
        strData = Space$(LOF(intI))
        Get #1, , strData
        Close #1
        Me.Caption = "Removing IDE3 tag..."
        strAll = strAll & Mid$(strData, IDE3.strLength - -1)
        DoEvents
        strData = Empty
    Next I
    Me.Caption = "Preparing to write IDE3 tag..."
    With IDE3
        .Songname = txtTitle.Text
        .Artist = txtArtist.Text
        .Album = txtAlbum.Text
        .Year = txtYear.Text
        .Comment = txtComment.Text
        .Genre = txtGenre.Text
    End With
    On Error Resume Next
    Me.Caption = "Deleting old file (if any)..."
    Kill FileName1
    Me.Caption = "Reading IDE3 settings..."
    If LenB(IDE3.Songname) = 0 Then
        IDE3.Songname = FileName1
    End If
    If LenB(IDE3.Artist) = 0 Then
        IDE3.Artist = "Sjoerd MP3 Merger"
    End If
    If LenB(IDE3.Album) = 0 Then
        IDE3.Album = "No album data"
    End If
    If LenB(IDE3.Year) = 0 Then
        IDE3.Year = Year(Now)
    End If
    If LenB(IDE3.Comment) = 0 Then
        IDE3.Comment = "MP3 merged with Sjoerd MP3 Merger"
    End If
    If LenB(IDE3.Genre) = 0 Then
        IDE3.Genre = "No genre data"
    End If
    Me.Caption = "Writing IDE3 Tag..."
    WriteTag FileName1, IDE3.Songname, IDE3.Artist, IDE3.Album, IDE3.Year, IDE3.Comment, IDE3.Genre
    Me.Caption = "Merging MP3's..."
    Open FileName1 For Binary As #1
    Put #1, , strAll
    Close #1
    Me.Caption = "Done"
    MsgBox "Note:" & vbCrLf & "Not all merged MP3's will work with every media player" & vbCrLf & "This code only deletes an IDE 3 V.1 Tag (V.2 tag can be different size at a different place)" & vbCrLf & "That's why some players think a new MP3 has started and thus stop playing", vbInformation, "Sjoerd MP3 Merger"
    On Error GoTo 0

End Sub

Private Sub Form_Load()

    cDialog.Filter = "MP3 files (*.mp3) |*.mp3"
    cDialog.FilterIndex = 1

End Sub

Private Function WriteTag(ByVal strFileName As String, _
                          ByVal Songname As String, _
                          ByVal Artist As String, _
                          ByVal Album As String, _
                          ByVal strYear As String, _
                          ByVal Comment As String, _
                          ByVal Genre As Integer) As Long

  
  Dim mp3File As Integer
  Dim sn      As String * 30
  Dim com     As String * 30
  Dim art     As String * 30
  Dim alb     As String * 30
  Dim yr      As String * 4

    Me.Tag = "TAG"
    sn = Songname
    com = Comment
    art = Artist
    alb = Album
    yr = strYear
    mp3File = FreeFile
    Open strFileName For Binary Access Write As #mp3File
    Seek #mp3File, FileLen(strFileName) - 127
    Put #mp3File, , Me.Tag
    Put #mp3File, , sn
    Put #mp3File, , art
    Put #mp3File, , alb
    Put #mp3File, , yr
    Put #mp3File, , com
    Close #mp3File

End Function

':)Roja's VB Code Fixer V1.1.78 (8-2-2004 16:58:24) 3 + 151 = 154 Lines Thanks Ulli for inspiration and lots of code.

