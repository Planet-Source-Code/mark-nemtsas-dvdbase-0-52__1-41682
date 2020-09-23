VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "DVDBase - The Easy DVD Database"
   ClientHeight    =   5250
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9960
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   664
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSort 
      Caption         =   "Sort"
      Height          =   555
      Left            =   0
      TabIndex        =   85
      Top             =   30
      Width           =   3225
      Begin VB.OptionButton optRatingSort 
         Caption         =   "Rating"
         Height          =   285
         Left            =   2400
         TabIndex        =   89
         ToolTipText     =   "Click to see a list by rating"
         Top             =   210
         Width           =   795
      End
      Begin VB.OptionButton optRegionSort 
         Caption         =   "Region"
         Height          =   285
         Left            =   1590
         TabIndex        =   88
         ToolTipText     =   "Click to view a list sorted by region"
         Top             =   210
         Width           =   825
      End
      Begin VB.OptionButton optGenreSort 
         Caption         =   "Genre"
         Height          =   285
         Left            =   870
         TabIndex        =   87
         ToolTipText     =   "Click to see a list sorted by genre"
         Top             =   210
         Width           =   795
      End
      Begin VB.OptionButton optNoSort 
         Caption         =   "None"
         Height          =   285
         Left            =   150
         TabIndex        =   86
         ToolTipText     =   "Click to view an unsorted DVD list"
         Top             =   210
         Width           =   795
      End
   End
   Begin MSComctlLib.TreeView treDiscs 
      Height          =   4095
      Left            =   30
      TabIndex        =   84
      Top             =   600
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   7223
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ilsTreeView"
      Appearance      =   1
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      Height          =   4515
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   6375
      Begin VB.CommandButton btnAddDVD 
         Caption         =   "Add DVD"
         Height          =   315
         Left            =   4470
         TabIndex        =   83
         ToolTipText     =   "Click to add this DVD"
         Top             =   180
         Width           =   945
      End
      Begin VB.ComboBox cboUserReview 
         Height          =   315
         ItemData        =   "frmMain.frx":0442
         Left            =   2550
         List            =   "frmMain.frx":0458
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   870
         Width           =   315
      End
      Begin VB.PictureBox picStar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   4
         Left            =   2220
         Picture         =   "frmMain.frx":046E
         ScaleHeight     =   195
         ScaleWidth      =   240
         TabIndex        =   81
         Top             =   900
         Width           =   240
      End
      Begin VB.PictureBox picStar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   3
         Left            =   2010
         Picture         =   "frmMain.frx":0570
         ScaleHeight     =   195
         ScaleWidth      =   240
         TabIndex        =   80
         Top             =   900
         Width           =   240
      End
      Begin VB.PictureBox picStar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   2
         Left            =   1800
         Picture         =   "frmMain.frx":0672
         ScaleHeight     =   195
         ScaleWidth      =   240
         TabIndex        =   79
         Top             =   900
         Width           =   240
      End
      Begin VB.PictureBox picStar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   1
         Left            =   1590
         Picture         =   "frmMain.frx":0774
         ScaleHeight     =   195
         ScaleWidth      =   240
         TabIndex        =   78
         Top             =   900
         Width           =   240
      End
      Begin VB.PictureBox picStar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   0
         Left            =   1380
         Picture         =   "frmMain.frx":0876
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   77
         Top             =   900
         Width           =   230
      End
      Begin VB.TextBox txtDirector 
         Height          =   315
         Left            =   4500
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtDatePurchased 
         Height          =   315
         Left            =   1620
         TabIndex        =   13
         Text            =   "66/66/66"
         Top             =   3720
         Width           =   795
      End
      Begin VB.TextBox txtLocationPurchased 
         Height          =   315
         Left            =   1620
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox txtCost 
         Height          =   315
         Left            =   1620
         TabIndex        =   14
         Text            =   "$666.66"
         Top             =   4080
         Width           =   705
      End
      Begin VB.TextBox txtStudio 
         Height          =   315
         Left            =   4500
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtRunningTime 
         Height          =   315
         Left            =   4500
         TabIndex        =   17
         Text            =   "66:66:66"
         ToolTipText     =   "Should be of the form hh:mm:ss, anything else will be rejected"
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox cboCurrentLocation 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2940
         Width           =   1815
      End
      Begin VB.ComboBox cboCaseType 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2580
         Width           =   1995
      End
      Begin VB.ComboBox cboRating 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2220
         Width           =   4935
      End
      Begin VB.ComboBox cboRegion 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1860
         Width           =   4935
      End
      Begin VB.TextBox txtDVDRelease 
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Text            =   "6666"
         Top             =   1500
         Width           =   495
      End
      Begin VB.TextBox txtMovieYear 
         Height          =   315
         Left            =   1380
         TabIndex        =   6
         Text            =   "6666"
         Top             =   1140
         Width           =   495
      End
      Begin VB.ComboBox cboGenre 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   540
         Width           =   1815
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   180
         Width           =   3015
      End
      Begin VB.Label Label16 
         Caption         =   "Director"
         Height          =   255
         Left            =   3420
         TabIndex        =   75
         Top             =   990
         Width           =   675
      End
      Begin VB.Label Label13 
         Caption         =   "Running Time"
         Height          =   255
         Left            =   3420
         TabIndex        =   66
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Studio"
         Height          =   255
         Left            =   3420
         TabIndex        =   65
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "Current Location"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   2970
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Case Type"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   2610
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   4110
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Location Purchased"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   3390
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Date Purchased"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   3750
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Rating"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   2250
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Region"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "DVD Release"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1530
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Movie Year"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "User Review"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Genre"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   210
         Width           =   1095
      End
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete DVD"
      Height          =   495
      Left            =   2130
      TabIndex        =   52
      Top             =   4740
      Width           =   1095
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "New DVD"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   4740
      Width           =   1095
   End
   Begin VB.Frame fraAudioVideo 
      Caption         =   "Audio / Video"
      Height          =   4455
      Left            =   3480
      TabIndex        =   2
      Top             =   540
      Width           =   6375
      Begin VB.Frame Frame6 
         Caption         =   "NTSC/PAL"
         Height          =   555
         Left            =   180
         TabIndex        =   74
         Top             =   1800
         Width           =   2955
         Begin VB.OptionButton optPAL 
            Caption         =   "PAL"
            Height          =   255
            Left            =   1620
            TabIndex        =   24
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton optNTSC 
            Caption         =   "NTSC"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Video Formats"
         Height          =   1515
         Left            =   180
         TabIndex        =   67
         Top             =   240
         Width           =   2955
         Begin VB.TextBox txtRatio 
            Height          =   315
            Left            =   2100
            TabIndex        =   22
            Text            =   "Text6"
            Top             =   270
            Width           =   555
         End
         Begin VB.CheckBox chk169 
            Caption         =   "16 x 9 Enhanced"
            Height          =   315
            Left            =   240
            TabIndex        =   21
            Top             =   1140
            Width           =   1635
         End
         Begin VB.CheckBox chkPanScan 
            Caption         =   "Pan & Scan"
            Height          =   315
            Left            =   240
            TabIndex        =   20
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox chkFullFrame 
            Caption         =   "Full Frame"
            Height          =   315
            Left            =   240
            TabIndex        =   19
            Top             =   540
            Width           =   1335
         End
         Begin VB.CheckBox chkWidescreen 
            Caption         =   "Widescreen"
            Height          =   315
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   " : 1"
            Height          =   255
            Left            =   2640
            TabIndex        =   71
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label14 
            Caption         =   "Ratio"
            Height          =   255
            Left            =   1620
            TabIndex        =   70
            Top             =   300
            Width           =   555
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Audio Formats"
         Height          =   3735
         Left            =   3240
         TabIndex        =   69
         Top             =   240
         Width           =   2955
         Begin VB.CheckBox chkAudioOther 
            Caption         =   "Other"
            Height          =   375
            Left            =   300
            TabIndex        =   40
            Top             =   2340
            Width           =   2235
         End
         Begin VB.CheckBox chkDolbySurround 
            Caption         =   "Dolby Surround"
            Height          =   375
            Left            =   300
            TabIndex        =   34
            Top             =   540
            Width           =   2235
         End
         Begin VB.CheckBox chkDolbyProLogic 
            Caption         =   "Dolby Pro-Logic"
            Height          =   375
            Left            =   300
            TabIndex        =   35
            Top             =   840
            Width           =   2235
         End
         Begin VB.CheckBox chkdd51 
            Caption         =   "Dolby Digital 5.1 (AC-3)"
            Height          =   375
            Left            =   300
            TabIndex        =   36
            Top             =   1140
            Width           =   2235
         End
         Begin VB.CheckBox chkDDEx 
            Caption         =   "Dolby Digital Surround EX"
            Height          =   375
            Left            =   300
            TabIndex        =   37
            Top             =   1440
            Width           =   2235
         End
         Begin VB.CheckBox chkDTS 
            Caption         =   "DTS"
            Height          =   375
            Left            =   300
            TabIndex        =   38
            Top             =   1740
            Width           =   2235
         End
         Begin VB.CheckBox chkSDDS 
            Caption         =   "Sony SDDS"
            Height          =   375
            Left            =   300
            TabIndex        =   39
            Top             =   2040
            Width           =   2235
         End
         Begin VB.CheckBox chkStereo 
            Caption         =   "Stereo"
            Height          =   375
            Left            =   300
            TabIndex        =   33
            Top             =   240
            Width           =   2235
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Subtitles"
         Height          =   1575
         Left            =   180
         TabIndex        =   68
         Top             =   2400
         Width           =   2955
         Begin VB.CheckBox chkFrench 
            Caption         =   "French"
            Height          =   315
            Left            =   240
            TabIndex        =   26
            Top             =   600
            Width           =   915
         End
         Begin VB.CheckBox chkGerman 
            Caption         =   "German"
            Height          =   315
            Left            =   240
            TabIndex        =   27
            Top             =   900
            Width           =   915
         End
         Begin VB.CheckBox chkSpanish 
            Caption         =   "Spanish"
            Height          =   315
            Left            =   240
            TabIndex        =   28
            Top             =   1200
            Width           =   915
         End
         Begin VB.CheckBox chkPortugese 
            Caption         =   "Portugese"
            Height          =   315
            Left            =   1620
            TabIndex        =   29
            Top             =   300
            Width           =   1035
         End
         Begin VB.CheckBox chkJapanese 
            Caption         =   "Japanese"
            Height          =   315
            Left            =   1620
            TabIndex        =   30
            Top             =   600
            Width           =   1035
         End
         Begin VB.CheckBox chkChinese 
            Caption         =   "Chinese"
            Height          =   315
            Left            =   1620
            TabIndex        =   31
            Top             =   900
            Width           =   915
         End
         Begin VB.CheckBox chkSubTitleOther 
            Caption         =   "Other"
            Height          =   315
            Left            =   1620
            TabIndex        =   32
            Top             =   1200
            Width           =   915
         End
         Begin VB.CheckBox chkEnglish 
            Caption         =   "English"
            Height          =   315
            Left            =   240
            TabIndex        =   25
            Top             =   300
            Width           =   915
         End
      End
   End
   Begin VB.Frame fraFeatures 
      Caption         =   "Features"
      Height          =   4455
      Left            =   3480
      TabIndex        =   3
      Top             =   540
      Width           =   6375
      Begin VB.Frame Frame5 
         Caption         =   "Disc Format"
         Height          =   3135
         Left            =   3060
         TabIndex        =   73
         Top             =   240
         Width           =   2115
         Begin VB.OptionButton optDualSided 
            Caption         =   "Dual-Sided"
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   660
            Width           =   1335
         End
         Begin VB.OptionButton optFlipper 
            Caption         =   "Flipper"
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton optDualLayer 
            Caption         =   "Dual Layer"
            Height          =   195
            Left            =   240
            TabIndex        =   49
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Disc Extras"
         Height          =   3135
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   2835
         Begin VB.CheckBox chkAnimatedMenus 
            Caption         =   "Animated Menus"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   600
            Width           =   1695
         End
         Begin VB.CheckBox chkMakingOf 
            Caption         =   """Making Of"" Documentary"
            Height          =   195
            Left            =   240
            TabIndex        =   43
            Top             =   900
            Width           =   2475
         End
         Begin VB.CheckBox chkBios 
            Caption         =   "Star Bios"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   2100
            Width           =   1035
         End
         Begin VB.CheckBox chkDeletedScenes 
            Caption         =   "Deleted Scenes"
            Height          =   195
            Left            =   240
            TabIndex        =   45
            Top             =   1500
            Width           =   1755
         End
         Begin VB.CheckBox chkCommentary 
            Caption         =   "Commentary"
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   1200
            Width           =   2235
         End
         Begin VB.CheckBox chkDVDROM 
            Caption         =   "DVD-ROM Content"
            Height          =   195
            Left            =   240
            TabIndex        =   48
            Top             =   2400
            Width           =   2295
         End
         Begin VB.CheckBox chkTheatricalTrailer 
            Caption         =   "Theatrical Trailer"
            Height          =   195
            Left            =   240
            TabIndex        =   46
            Top             =   1800
            Width           =   1875
         End
         Begin VB.CheckBox chkSceneAccess 
            Caption         =   "Scene Access"
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   300
            Width           =   1635
         End
      End
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   5205
      Left            =   3270
      TabIndex        =   76
      Top             =   30
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   9181
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Object.Tag             =   "General"
            Object.ToolTipText     =   "Click to see the general properties of this DVD"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Audio / Video"
            Object.Tag             =   "AudioVideo"
            Object.ToolTipText     =   "Click to see the Audio and Video properties of this DVD"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Features"
            Object.Tag             =   "Features"
            Object.ToolTipText     =   "Click to see the features of this DVD and see the disc type"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsTreeView 
      Left            =   1320
      Top             =   4650
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0978
            Key             =   "Disc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DCA
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":121C
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDragNode As Node
Private blnInDrag As Boolean


Private Sub btnAddDVD_Click()
  If intFormAction = ADD_NEW Then
    If Len(Me.txtTitle) > 0 Then
      If vbYes = MsgBox("Are you sure you want to add this DVD?", vbYesNo, "Confirm Add DVD") Then
        dvdCurrent.addDVD Me.txtTitle
        If Me.optNoSort = True Then
          globalCode.fillDVDTreeView Me.treDiscs, , discCode.getLatestDVD
        End If
        If Me.optGenreSort = True Then
          globalCode.fillDVDTreeView Me.treDiscs, "Genre", discCode.getLatestDVD
        End If
        If Me.optRegionSort = True Then
          globalCode.fillDVDTreeView Me.treDiscs, "Region", discCode.getLatestDVD
        End If
        If Me.optRatingSort = True Then
          globalCode.fillDVDTreeView Me.treDiscs, "Rating", discCode.getLatestDVD
        End If
        Me.Caption = strCaption & " : " & dvdCurrent.strTitle
        Me.btnAddDVD.Visible = False
      End If
    End If
  End If
End Sub

Private Sub btnDelete_Click()
  If checkSelected(Me.treDiscs) = -1 Then Exit Sub
  If vbYes = MsgBox("Are you sure you want to delete this DVD?", vbYesNo, "Deletion Warning") Then
    dvdCurrent.deleteDVD
    Me.Caption = strCaption
    If Me.optNoSort = True Then
      globalCode.fillDVDTreeView Me.treDiscs
    End If
    If Me.optGenreSort = True Then
      globalCode.fillDVDTreeView Me.treDiscs, "Genre"
    End If
    If Me.optRegionSort = True Then
      globalCode.fillDVDTreeView Me.treDiscs, "Region"
    End If
    If Me.optRatingSort = True Then
      globalCode.fillDVDTreeView Me.treDiscs, "Rating"
    End If
    discCode.resetAllFields
    discCode.disableDiscDisplay
  End If

End Sub

Private Sub btnNew_Click()
  intFormAction = ADD_NEW
  dvdCurrent.closeDVD
  If Not Me.treDiscs.SelectedItem Is Nothing Then
    Me.treDiscs.Nodes(Me.treDiscs.SelectedItem.Index).Selected = False
  End If
  discCode.enableDiscDisplay
  discCode.resetAllFields
  discCode.enableAddDiscDisplay
  Me.btnAddDVD.Visible = True
End Sub

Private Sub cboCaseType_Click()
  If Me.cboCaseType.ListIndex <> -1 Then
    dvdCurrent.lngCaseType = lngCaseTypeArray(Me.cboCaseType.ListIndex + 1)
  Else
    dvdCurrent.lngCaseType = 0
  End If

End Sub

Private Sub cboCurrentLocation_Click()
  If Me.cboCurrentLocation.ListIndex <> -1 Then
    dvdCurrent.lngCurrentLocation = lngCurrentLocationArray(Me.cboCurrentLocation.ListIndex + 1)
  Else
    dvdCurrent.lngCurrentLocation = 0
  End If

End Sub

Private Sub cboGenre_Click()
  If Me.cboGenre.ListIndex <> -1 Then
    dvdCurrent.lngGenre = lngGenreArray(Me.cboGenre.ListIndex + 1)
    updateTreeView "Genre", Me.cboGenre
  Else
    dvdCurrent.lngGenre = 0
  End If
End Sub

Private Sub cboRating_Click()
  If Me.cboRating.ListIndex <> -1 Then
    dvdCurrent.lngRating = lngRatingArray(Me.cboRating.ListIndex + 1)
    updateTreeView "Rating", Me.cboRating
  Else
    dvdCurrent.lngRating = 0
  End If
End Sub

Private Sub cboRegion_Click()
  If Me.cboRegion.ListIndex <> -1 Then
    dvdCurrent.lngRegion = lngRegionArray(Me.cboRegion.ListIndex + 1)
    updateTreeView "Region", Me.cboRegion
  Else
    dvdCurrent.lngRegion = 0
  End If
End Sub

Private Sub cboUserReview_Click()
  Dim intStars As Integer, intLoop As Integer
  intStars = Me.cboUserReview
  For intLoop = 0 To 4
    Me.picStar(intLoop).Visible = False
  Next intLoop
  For intLoop = 0 To intStars - 1
    Me.picStar(intLoop).Visible = True
  Next intLoop
  dvdCurrent.bytUserReview = CByte(intStars)
End Sub

Private Sub chk169_Validate(Cancel As Boolean)
  dvdCurrent.bln169 = Me.chk169
End Sub

Private Sub chkAnimatedMenus_Validate(Cancel As Boolean)
  dvdCurrent.blnAnimatedMenus = Me.chkAnimatedMenus
End Sub

Private Sub chkAudioOther_Validate(Cancel As Boolean)
  dvdCurrent.blnAudioOther = Me.chkAudioOther
End Sub

Private Sub chkBios_Validate(Cancel As Boolean)
  dvdCurrent.blnStarBios = Me.chkBios
End Sub

Private Sub chkChinese_Validate(Cancel As Boolean)
  dvdCurrent.blnChinese = Me.chkChinese
End Sub

Private Sub chkCommentary_Validate(Cancel As Boolean)
  dvdCurrent.blnCommentary = Me.chkCommentary
End Sub

Private Sub chkdd51_Validate(Cancel As Boolean)
  dvdCurrent.blnDD51 = Me.chkdd51
End Sub

Private Sub chkDDEx_Validate(Cancel As Boolean)
  dvdCurrent.blnDDEx = Me.chkDDEx
End Sub

Private Sub chkDeletedScenes_Validate(Cancel As Boolean)
  dvdCurrent.blnDeletedScenes = Me.chkDeletedScenes
End Sub

Private Sub chkDolbyProLogic_Validate(Cancel As Boolean)
  dvdCurrent.blnDolbyProLogic = Me.chkDolbyProLogic
End Sub

Private Sub chkDolbySurround_Validate(Cancel As Boolean)
  dvdCurrent.blnDolbySurround = Me.chkDolbySurround
End Sub

Private Sub chkDTS_Validate(Cancel As Boolean)
  dvdCurrent.blnDTS = Me.chkDTS
End Sub

Private Sub chkDVDROM_Validate(Cancel As Boolean)
  dvdCurrent.blnDVDRom = Me.chkDVDROM
End Sub

Private Sub chkEnglish_Validate(Cancel As Boolean)
  dvdCurrent.blnEnglish = Me.chkEnglish
End Sub

Private Sub chkFrench_Validate(Cancel As Boolean)
  dvdCurrent.blnFrench = Me.chkFrench
End Sub

Private Sub chkFullFrame_Validate(Cancel As Boolean)
  dvdCurrent.blnFullFrame = Me.chkFullFrame
End Sub

Private Sub chkGerman_Validate(Cancel As Boolean)
  dvdCurrent.blnGerman = Me.chkGerman
End Sub

Private Sub chkJapanese_Validate(Cancel As Boolean)
  dvdCurrent.blnJapanese = Me.chkJapanese
End Sub

Private Sub chkMakingOf_Validate(Cancel As Boolean)
  dvdCurrent.blnMakingOf = Me.chkMakingOf
End Sub

Private Sub chkPanScan_Validate(Cancel As Boolean)
  dvdCurrent.blnPanScan = Me.chkPanScan
End Sub

Private Sub chkPortugese_Validate(Cancel As Boolean)
  dvdCurrent.blnPortugese = Me.chkPortugese
End Sub

Private Sub chkSceneAccess_Validate(Cancel As Boolean)
  dvdCurrent.blnSceneAccess = Me.chkSceneAccess
End Sub

Private Sub chkSDDS_Validate(Cancel As Boolean)
  dvdCurrent.blnSDDS = Me.chkSDDS
End Sub

Private Sub chkSpanish_Validate(Cancel As Boolean)
  dvdCurrent.blnSpanish = Me.chkSpanish
End Sub

Private Sub chkStereo_Validate(Cancel As Boolean)
  dvdCurrent.blnStereo = Me.chkStereo
End Sub

Private Sub chkSubTitleOther_Validate(Cancel As Boolean)
  dvdCurrent.blnSubtitleOther = Me.chkSubTitleOther
End Sub

Private Sub chkTheatricalTrailer_Validate(Cancel As Boolean)
  dvdCurrent.blnTheatricalTrailer = Me.chkTheatricalTrailer
End Sub

Private Sub chkWidescreen_Validate(Cancel As Boolean)
  dvdCurrent.blnWidescreen = Me.chkWidescreen
End Sub

Private Sub Form_Load()
  globalCode.initialise
  Me.optNoSort.Value = True
  
  Me.fraGeneral.Visible = True
  Me.fraAudioVideo.Visible = False
  Me.fraFeatures.Visible = False
  '
  'Fill combos
  '
  globalCode.fillSelectCombo Me.cboGenre.Name
  globalCode.fillSelectCombo Me.cboRegion.Name
  globalCode.fillSelectCombo Me.cboRating.Name
  globalCode.fillSelectCombo Me.cboCaseType.Name
  globalCode.fillSelectCombo Me.cboCurrentLocation.Name
  
  discCode.resetAllFields
  discCode.disableDiscDisplay
  
  Me.btnAddDVD.Visible = False
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show
End Sub

Private Sub optDualLayer_Click()
  If Me.optDualLayer = True Then dvdCurrent.bytDiscFormat = 1
  If Me.optDualSided = True Then dvdCurrent.bytDiscFormat = 2
  If Me.optFlipper = True Then dvdCurrent.bytDiscFormat = 3
End Sub

Private Sub optDualSided_Click()
  If Me.optDualLayer = True Then dvdCurrent.bytDiscFormat = 1
  If Me.optDualSided = True Then dvdCurrent.bytDiscFormat = 2
  If Me.optFlipper = True Then dvdCurrent.bytDiscFormat = 3
End Sub

Private Sub optFlipper_Click()
  If Me.optDualLayer = True Then dvdCurrent.bytDiscFormat = 1
  If Me.optDualSided = True Then dvdCurrent.bytDiscFormat = 2
  If Me.optFlipper = True Then dvdCurrent.bytDiscFormat = 3
End Sub

Private Sub optGenreSort_Click()
  If checkSelected(Me.treDiscs) <> -1 Then
    globalCode.fillDVDTreeView Me.treDiscs, "Genre", checkSelected(Me.treDiscs)
  Else
    globalCode.fillDVDTreeView Me.treDiscs, "Genre"
  End If
End Sub

Private Sub optNoSort_Click()
  If checkSelected(Me.treDiscs) <> -1 Then
    globalCode.fillDVDTreeView Me.treDiscs, , checkSelected(Me.treDiscs)
  Else
    globalCode.fillDVDTreeView Me.treDiscs
  End If

End Sub

Private Sub optNTSC_Click()
  dvdCurrent.blnNTSCPAL = Me.optNTSC
End Sub

Private Sub optPAL_Click()
  dvdCurrent.blnNTSCPAL = Me.optNTSC
End Sub

Private Sub optRatingSort_Click()
  If checkSelected(Me.treDiscs) <> -1 Then
    globalCode.fillDVDTreeView Me.treDiscs, "Rating", checkSelected(Me.treDiscs)
  Else
    globalCode.fillDVDTreeView Me.treDiscs, "Rating"
  End If
End Sub

Private Sub optRegionSort_Click()
 
  If checkSelected(Me.treDiscs) <> -1 Then
    globalCode.fillDVDTreeView Me.treDiscs, "Region", checkSelected(Me.treDiscs)
  Else
    globalCode.fillDVDTreeView Me.treDiscs, "Region"
  End If
End Sub

Private Sub tabMain_Click()
  Select Case Me.tabMain.SelectedItem.Tag
    Case "General"
      Me.fraGeneral.Visible = True
      Me.fraAudioVideo.Visible = False
      Me.fraFeatures.Visible = False
      dvdCurrent.displayDVD
    Case "AudioVideo"
      Me.fraGeneral.Visible = False
      Me.fraAudioVideo.Visible = True
      Me.fraFeatures.Visible = False
      dvdCurrent.displayDVD
    Case "Features"
      Me.fraGeneral.Visible = False
      Me.fraAudioVideo.Visible = False
      Me.fraFeatures.Visible = True
      dvdCurrent.displayDVD
  End Select
End Sub


Private Sub treDiscs_Click()
  Dim nNode As Node
  Dim lngIndex As Long
  
  Set nNode = Me.treDiscs.SelectedItem
  If Left(nNode.Key, 5) = "Node_" Then
    lngIndex = CLng(Right(nNode.Key, Len(nNode.Key) - Len("Node_")))
    If Me.txtDirector.Enabled = False Then discCode.enableDiscDisplay
    intFormAction = EDIT
    Me.btnAddDVD.Visible = False
    dvdCurrent.fillDVD lngIndex
    dvdCurrent.displayDVD
    Me.Caption = strCaption & " : " & dvdCurrent.strTitle
  Else
    Exit Sub
  End If
End Sub

Private Sub treDiscs_Collapse(ByVal Node As MSComctlLib.Node)
  If Node.Parent Is Nothing Then Exit Sub
  If Node.Parent.Key = "Collection" Then
    Node.Image = "Closed"
  End If
End Sub


Private Sub treDiscs_DragDrop(Source As Control, x As Single, y As Single)
  Set Me.treDiscs.DropHighlight = Me.treDiscs.HitTest(x, y)
  If Me.treDiscs.DropHighlight Is Nothing Then
    blnInDrag = False
    Set Me.treDiscs.DropHighlight = Nothing
    Set mDragNode = Nothing
    Exit Sub
  End If
  If Left(Me.treDiscs.DropHighlight.Key, 5) = "Node_" Or Me.treDiscs.DropHighlight.Key = "Collection" Or Me.treDiscs.DropHighlight = mDragNode Or Left(Me.treDiscs.DropHighlight, 3) = "No " Then
    Debug.Print "Invalid Drop Location"
  Else
    If Me.optGenreSort = True Then
      discCode.updateTreeViewDragDrop "Genre", Me.cboGenre, mDragNode, Me.treDiscs.DropHighlight
    End If
    If Me.optRegionSort = True Then
      discCode.updateTreeViewDragDrop "Region", Me.cboRegion, mDragNode, Me.treDiscs.DropHighlight
    End If
    If Me.optRatingSort = True Then
      discCode.updateTreeViewDragDrop "Rating", Me.cboRating, mDragNode, Me.treDiscs.DropHighlight
    End If
  End If
  blnInDrag = False
  Set Me.treDiscs.DropHighlight = Nothing
  Set mDragNode = Nothing
End Sub

Private Sub treDiscs_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If blnInDrag = True And Not mDragNode Is Nothing Then
        ' Set DropHighlight to the mouse's coordinates.
        Set Me.treDiscs.DropHighlight = Me.treDiscs.HitTest(x, y)
    End If
End Sub

Private Sub treDiscs_Expand(ByVal Node As MSComctlLib.Node)
  If Node.Parent Is Nothing Then Exit Sub
  If Node.Parent.Key = "Collection" Then
    Node.Image = "Open"
  End If
End Sub


Private Sub treDiscs_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button <> vbLeftButton Then Exit Sub
  If Me.treDiscs.HitTest(x, y) Is Nothing Then Exit Sub
  Set Me.treDiscs.SelectedItem = Me.treDiscs.HitTest(x, y)
  treDiscs_Click
  Set Me.treDiscs.DropHighlight = Me.treDiscs.HitTest(x, y)
  'Make sure we are over a Node
  If Not Me.treDiscs.DropHighlight Is Nothing Then
    'Set the Node we are on to be the selected Node
    'if we don't do this it will not be the selected node
    'until we finish clicking on the Node
    If Left(Me.treDiscs.SelectedItem.Key, 5) = "Node_" Then
      Set mDragNode = Me.treDiscs.SelectedItem ' Set the item being dragged.
      Exit Sub
    End If
  End If
  Set Me.treDiscs.DropHighlight = Nothing
  
End Sub

Private Sub treDiscs_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton And Not mDragNode Is Nothing Then ' Signal a Drag operation.
      blnInDrag = True ' Set the flag to true.
      ' Set the drag icon with the CreateDragImage method.
      Me.treDiscs.DragIcon = Me.treDiscs.SelectedItem.CreateDragImage
      Me.treDiscs.Drag vbBeginDrag ' Drag operation.
    End If
End Sub

Private Sub txtCost_Validate(Cancel As Boolean)
  If Len(Me.txtCost) = 0 Then Exit Sub
  If IsNumeric(Me.txtCost) = True Then
    dvdCurrent.curCost = Me.txtCost
  Else
    MsgBox "Cost must be a number"
    Cancel = True
  End If
End Sub

Private Sub txtDatePurchased_Validate(Cancel As Boolean)
  If Len(Me.txtDatePurchased) = 0 Then Exit Sub
  If IsDate(Me.txtDatePurchased) = True Then
    dvdCurrent.datDatePurchased = Me.txtDatePurchased
  Else
    MsgBox "Date purchased must be a date"
    Cancel = True
  End If
End Sub

Private Sub txtDirector_Validate(Cancel As Boolean)
  dvdCurrent.strDirector = Me.txtDirector
End Sub

Private Sub txtDVDRelease_Validate(Cancel As Boolean)
  If Len(Me.txtDVDRelease) = 0 Then Exit Sub
  If IsNumeric(Me.txtDVDRelease) = True And Len(Me.txtDVDRelease) = 4 Then
    dvdCurrent.datDVDRelease = "1/1/" & Me.txtDVDRelease
  Else
    MsgBox "DVD release must be a year"
    Cancel = True
  End If
End Sub

Private Sub txtLocationPurchased_Validate(Cancel As Boolean)
  dvdCurrent.strLocationPurchased = Me.txtLocationPurchased
End Sub


Private Sub txtMovieYear_Validate(Cancel As Boolean)
  If Len(Me.txtMovieYear) = 0 Then Exit Sub
  If IsNumeric(Me.txtMovieYear) = True And Len(Me.txtMovieYear) = 4 Then
    dvdCurrent.datMovieYear = "1/1/" & Me.txtMovieYear
  Else
    MsgBox "Movie year must be a year"
    Cancel = True
  End If
End Sub

Private Sub txtRatio_Validate(Cancel As Boolean)
  If Len(Me.txtRatio) = 0 Then Exit Sub
  If IsNumeric(Me.txtRatio) = True Then
    dvdCurrent.dblRatio = Me.txtRatio
  Else
    MsgBox "Ratio must be a number"
    Cancel = True
  End If

End Sub

Private Sub txtRunningTime_Validate(Cancel As Boolean)
  If Len(Me.txtRunningTime) = 0 Then Exit Sub
  Dim intRunningTime As Integer
  intRunningTime = discCode.parseTime(Me.txtRunningTime)
  If intRunningTime <> -1 Then
    dvdCurrent.intRunningTime = intRunningTime
  Else
    MsgBox "Running time must be of form hh, hh:mm, or hh:mm:ss, and total seconds must be less than 65000 seconds " & Chr(10) & _
    "which is 18 hours.  Show me a DVD with a running time more than this and I'll make this field a long."
    Cancel = True
  End If
End Sub

Private Sub txtStudio_Validate(Cancel As Boolean)
  dvdCurrent.strStudio = Me.txtStudio
End Sub

Private Sub txtTitle_Change()
  If intFormAction = ADD_NEW And Len(Me.txtTitle) > 0 And Len(Me.txtTitle) < 150 Then Me.btnAddDVD.Enabled = True
End Sub

Private Sub txtTitle_Validate(Cancel As Boolean)
  If intFormAction = ADD_NEW Then
    If Len(Me.txtTitle) = 0 Then
      MsgBox "DVD Title must have 1 or more characters"
      Cancel = True
    End If
  End If
  If intFormAction = EDIT Then
    If Len(Me.txtTitle) = 0 And Len(Me.txtTitle) < 150 Then
      MsgBox "Cannot update, DVD Title must have more than 1 character and less than 150 characters"
      Cancel = True
    Else
      dvdCurrent.strTitle = Me.txtTitle
    End If
  End If
End Sub

