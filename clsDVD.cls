VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDVD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarlngID As Long 'local copy
Private mvarstrTitle As String 'local copy
Private mvarlngGenre As Long 'local copy
Private mvarbytUserReview As Byte 'local copy
Private mvardatMovieYear As Date 'local copy
Private mvardatDVDRelease As Date 'local copy
Private mvarlngRegion As Long 'local copy
Private mvarlngRating As Long 'local copy
Private mvarlngCaseType As Long 'local copy
Private mvarlngCurrentLocation As Long 'local copy
Private mvarstrLocationPurchased As String 'local copy
Private mvardatDatePurchased As Date 'local copy
Private mvarcurCost As Currency 'local copy
Private mvarstrStudio As String 'local copy
Private mvarstrDirector As String 'local copy
Private mvarblnWidescreen As Boolean 'local copy
Private mvarblnFullFrame As Boolean 'local copy
Private mvarblnPanScan As Boolean 'local copy
Private mvarbln169 As Boolean 'local copy
Private mvardblRatio As Double 'local copy
Private mvarintRunningTime As Integer
Private mvarblnNTSCPAL As Boolean 'local copy
Private mvarblnEnglish As Boolean 'local copy
Private mvarblnFrench As Boolean 'local copy
Private mvarblnGerman As Boolean 'local copy
Private mvarblnSpanish As Boolean 'local copy
Private mvarblnPortugese As Boolean 'local copy
Private mvarblnJapanese As Boolean 'local copy
Private mvarblnChinese As Boolean 'local copy
Private mvarblnSubtitleOther As Boolean 'local copy
Private mvarblnStereo As Boolean 'local copy
Private mvarblnDolbySurround As Boolean 'local copy
Private mvarblnDolbyProLogic As Boolean 'local copy
Private mvarblnDD51 As Boolean 'local copy
Private mvarblnDDEx As Boolean 'local copy
Private mvarblnDTS As Boolean 'local copy
Private mvarblnSDDS As Boolean 'local copy
Private mvarblnAudioOther As Boolean 'local copy
Private mvarblnSceneAccess As Boolean 'local copy
Private mvarblnAnimatedMenus As Boolean 'local copy
Private mvarblnMakingOf As Boolean 'local copy
Private mvarblnCommentary As Boolean 'local copy
Private mvarblnDeletedScenes As Boolean 'local copy
Private mvarblnTheatricalTrailer As Boolean 'local copy
Private mvarblnStarBios As Boolean 'local copy
Private mvarblnDVDRom As Boolean 'local copy
Private mvarbytDiscFormat As Byte 'local copy
Private rstDVD As New ADODB.Recordset 'local copy

Public Sub deleteDVD()
  If rstDVD.State = adStateClosed Then Exit Sub
  rstDVD.DELETE
  rstDVD.Update
  rstDVD.Close
  resetDVD
End Sub

Public Sub closeDVD()
  If rstDVD.State = adStateOpen Then
    rstDVD.Update
    rstDVD.Close
    resetDVD
  End If
End Sub

Public Sub addDVD(strTitle As String)
  If rstDVD.State = adStateOpen Then
    rstDVD.Update
    rstDVD.Close
  End If
  cmdSelectDVDByID.Parameters(0) = -1
  Set rstDVD = returnRS(cmdSelectDVDByID)
  If rstDVD.EOF = True Then
    With rstDVD
      .AddNew
        ![strTitle] = strTitle
        ![datMovieYear] = 0
        ![datDVDRelease] = 0
        ![strLocationPurchased] = ""
        ![datDatePurchased] = 0
        ![strStudio] = ""
        ![strDirector] = ""
      .Update
    End With
  Else
    MsgBox "Error: Your database seems to be corrupted, a record was found for ID: -1 "
    rstDVD.Close
  End If
End Sub

Public Sub fillDVD(lngID As Long)
  If lngID <= 0 Then Exit Sub
  cmdSelectDVDByID.Parameters(0) = lngID
  If rstDVD.State = adStateOpen Then rstDVD.Close
  Set rstDVD = returnRS(cmdSelectDVDByID)
  If rstDVD.EOF <> True Then
    mvarlngID = rstDVD![lngID]
    mvarstrTitle = rstDVD![strTitle]
    mvarlngGenre = rstDVD![lngGenre]
    mvarbytUserReview = rstDVD![bytUserReview]
    mvardatMovieYear = rstDVD![datMovieYear]
    mvardatDVDRelease = rstDVD![datDVDRelease]
    mvarlngRegion = rstDVD![lngRegion]
    mvarlngRating = rstDVD![lngRating]
    mvarlngCaseType = rstDVD![lngCaseType]
    mvarlngCurrentLocation = rstDVD![lngCurrentLocation]
    mvarstrLocationPurchased = rstDVD![strLocationPurchased]
    mvardatDatePurchased = rstDVD![datDatePurchased]
    mvarcurCost = rstDVD![curCost]
    mvarstrStudio = rstDVD![strStudio]
    mvarstrDirector = rstDVD![strDirector]
    mvarblnWidescreen = rstDVD![blnWidescreen]
    mvarblnFullFrame = rstDVD![blnFullFrame]
    mvarblnPanScan = rstDVD![blnPanScan]
    mvarbln169 = rstDVD![bln169]
    mvardblRatio = rstDVD![dblRatio]
    mvarintRunningTime = rstDVD![intRunningTime]
    mvarblnNTSCPAL = rstDVD![blnNTSCPAL]
    mvarblnEnglish = rstDVD![blnEnglish]
    mvarblnFrench = rstDVD![blnFrench]
    mvarblnGerman = rstDVD![blnGerman]
    mvarblnSpanish = rstDVD![blnSpanish]
    mvarblnPortugese = rstDVD![blnPortugese]
    mvarblnJapanese = rstDVD![blnJapanese]
    mvarblnChinese = rstDVD![blnChinese]
    mvarblnSubtitleOther = rstDVD![blnSubtitleOther]
    mvarblnStereo = rstDVD![blnStereo]
    mvarblnDolbySurround = rstDVD![blnDolbySurround]
    mvarblnDolbyProLogic = rstDVD![blnDolbyProLogic]
    mvarblnDD51 = rstDVD![blnDD51]
    mvarblnDDEx = rstDVD![blnDDEx]
    mvarblnDTS = rstDVD![blnDTS]
    mvarblnSDDS = rstDVD![blnSDDS]
    mvarblnAudioOther = rstDVD![blnAudioOther]
    mvarblnSceneAccess = rstDVD![blnSceneAccess]
    mvarblnAnimatedMenus = rstDVD![blnAnimatedMenus]
    mvarblnMakingOf = rstDVD![blnMakingOf]
    mvarblnCommentary = rstDVD![blnCommentary]
    mvarblnDeletedScenes = rstDVD![blnDeletedScenes]
    mvarblnTheatricalTrailer = rstDVD![blnTheatricalTrailer]
    mvarblnStarBios = rstDVD![blnStarBios]
    mvarblnDVDRom = rstDVD![blnDVDRom]
    mvarbytDiscFormat = rstDVD![bytDiscFormat]
  Else
    MsgBox "Error: Your database seems to be corrupted, a record could not be found for ID: " & lngID
    rstDVD.Close
  End If
End Sub

Public Sub resetDVD()
    mvarlngID = 0
    mvarstrTitle = ""
    mvarlngGenre = 0
    mvarbytUserReview = 0
    mvardatMovieYear = 0
    mvardatDVDRelease = 0
    mvarlngRegion = 0
    mvarlngRating = 0
    mvarlngCaseType = 0
    mvarlngCurrentLocation = 0
    mvarstrLocationPurchased = ""
    mvardatDatePurchased = 0
    mvarcurCost = 0
    mvarstrStudio = ""
    mvarstrDirector = ""
    mvardblRatio = 0
    mvarintRunningTime = 0
    mvarbytDiscFormat = 0
End Sub

Public Sub displayDVD()
  Dim intComboIndex As Integer
  Dim intLoop As Integer
  Dim strHours As String, strMinutes As String, strSeconds As String
  Dim lngStart As Long
  
  lngStart = GetTickCount
  If rstDVD.State = adStateClosed Then Exit Sub
  
  Select Case frmMain.tabMain.SelectedItem.Tag
    Case "General"
  frmMain.txtTitle = mvarstrTitle
  If mvarlngGenre > 0 Then
    intComboIndex = discCode.returnComboLocation(lngGenreArray, intGenreArraySize, mvarlngGenre)
  Else
    intComboIndex = -1
  End If
  If intComboIndex < 0 Then
    'rstDVD![lngGenre] = 0
    'rstDVD.Update
    frmMain.cboGenre.ListIndex = -1
  Else
    frmMain.cboGenre.ListIndex = intComboIndex - 1
  End If
  For intLoop = 0 To 4
    frmMain.picStar(intLoop).Visible = False
  Next intLoop
  If mvarbytUserReview > 0 Then
    For intLoop = 0 To mvarbytUserReview - 1
      frmMain.picStar(intLoop).Visible = True
    Next intLoop
  End If
  If mvardatMovieYear > 0 Then
    frmMain.txtMovieYear = Year(mvardatMovieYear)
  Else
    frmMain.txtMovieYear = ""
  End If
  If mvardatDVDRelease > 0 Then
    frmMain.txtDVDRelease = Year(mvardatDVDRelease)
  Else
    frmMain.txtDVDRelease = ""
  End If
  If mvarlngRegion > 0 Then
    intComboIndex = discCode.returnComboLocation(lngRegionArray, intRegionArraySize, mvarlngRegion)
  Else
    intComboIndex = -1
  End If
  If intComboIndex < 0 Then
    'rstDVD![lngRegion] = 0
    'rstDVD.Update
    frmMain.cboRegion.ListIndex = -1
  Else
    frmMain.cboRegion.ListIndex = intComboIndex - 1
  End If
  If mvarlngRating > 0 Then
    intComboIndex = discCode.returnComboLocation(lngRatingArray, intRatingArraySize, mvarlngRating)
  Else
    intComboIndex = -1
  End If
  If intComboIndex < 0 Then
    'rstDVD![lngRating] = 0
    'rstDVD.Update
    frmMain.cboRating.ListIndex = -1
  Else
    frmMain.cboRating.ListIndex = intComboIndex - 1
  End If
  If mvarlngCaseType > 0 Then
    intComboIndex = discCode.returnComboLocation(lngCaseTypeArray, intCaseTypeArraySize, mvarlngCaseType)
  Else
    intComboIndex = -1
  End If
  If intComboIndex < 0 Then
    'rstDVD![lngCaseType] = 0
    'rstDVD.Update
    frmMain.cboCaseType.ListIndex = -1
  Else
    frmMain.cboCaseType.ListIndex = intComboIndex - 1
  End If
  If mvarlngCurrentLocation > 0 Then
  intComboIndex = discCode.returnComboLocation(lngCurrentLocationArray, intCurrentLocationArraySize, mvarlngCurrentLocation)
  Else
    intComboIndex = -1
  End If
  If intComboIndex < 0 Then
    'rstDVD![lngCurrentLocation] = 0
    'rstDVD.Update
   frmMain.cboCurrentLocation.ListIndex = -1
  Else
    frmMain.cboCurrentLocation.ListIndex = intComboIndex - 1
  End If
  frmMain.txtLocationPurchased = mvarstrLocationPurchased
  If mvardatDatePurchased > 0 Then
    frmMain.txtDatePurchased = mvardatDatePurchased
  Else
    frmMain.txtDatePurchased = ""
  End If
  frmMain.txtCost = "$ " & mvarcurCost
  frmMain.txtStudio = mvarstrStudio
  frmMain.txtDirector = mvarstrDirector
  strHours = mvarintRunningTime \ 3600
  strMinutes = (mvarintRunningTime Mod 3600) \ 60
  If Len(strMinutes) = 1 Then strMinutes = "0" & strMinutes
  strSeconds = mvarintRunningTime Mod 60
  If Len(strSeconds) = 1 Then strSeconds = "0" & strSeconds
  frmMain.txtRunningTime = strHours & ":" & strMinutes & ":" & strSeconds
    Case "AudioVideo"
  frmMain.chkWidescreen = returnByte(mvarblnWidescreen)
  frmMain.chkFullFrame = returnByte(mvarblnFullFrame)
  frmMain.chkPanScan = returnByte(mvarblnPanScan)
  frmMain.chk169 = returnByte(mvarbln169)
  frmMain.txtRatio = mvardblRatio
  frmMain.optNTSC = mvarblnNTSCPAL
  frmMain.optPAL = Not (mvarblnNTSCPAL)
  frmMain.chkEnglish = returnByte(mvarblnEnglish)
  frmMain.chkFrench = returnByte(mvarblnFrench)
  frmMain.chkGerman = returnByte(mvarblnGerman)
  frmMain.chkSpanish = returnByte(mvarblnSpanish)
  frmMain.chkPortugese = returnByte(mvarblnPortugese)
  frmMain.chkJapanese = returnByte(mvarblnJapanese)
  frmMain.chkChinese = returnByte(mvarblnChinese)
  frmMain.chkSubTitleOther = returnByte(mvarblnSubtitleOther)
  frmMain.chkStereo = returnByte(mvarblnStereo)
  frmMain.chkDolbySurround = returnByte(mvarblnDolbySurround)
  frmMain.chkDolbyProLogic = returnByte(mvarblnDolbyProLogic)
  frmMain.chkdd51 = returnByte(mvarblnDD51)
  frmMain.chkDDEx = returnByte(mvarblnDDEx)
  frmMain.chkDTS = returnByte(mvarblnDTS)
  frmMain.chkSDDS = returnByte(mvarblnSDDS)
    Case "Features"
  frmMain.chkAudioOther = returnByte(mvarblnAudioOther)
  frmMain.chkSceneAccess = returnByte(mvarblnSceneAccess)
  frmMain.chkAnimatedMenus = returnByte(mvarblnAnimatedMenus)
  frmMain.chkMakingOf = returnByte(mvarblnMakingOf)
  frmMain.chkCommentary = returnByte(mvarblnCommentary)
  frmMain.chkDeletedScenes = returnByte(mvarblnDeletedScenes)
  frmMain.chkTheatricalTrailer = returnByte(mvarblnTheatricalTrailer)
  frmMain.chkBios = returnByte(mvarblnStarBios)
  frmMain.chkDVDROM = returnByte(mvarblnDVDRom)
  Select Case mvarbytDiscFormat
    Case 1
      frmMain.optDualLayer = True
      frmMain.optDualSided = False
      frmMain.optFlipper = False
    Case 2
      frmMain.optDualLayer = False
      frmMain.optDualSided = True
      frmMain.optFlipper = False
    Case 3
      frmMain.optDualLayer = False
      frmMain.optDualSided = False
      frmMain.optFlipper = True
  End Select
  End Select
  
  Debug.Print "Time: " & (GetTickCount - lngStart)
End Sub

Public Property Let bytDiscFormat(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bytDiscFormat = 5
  If rstDVD.State = adStateOpen Then
    mvarbytDiscFormat = vData
    rstDVD![bytDiscFormat] = mvarbytDiscFormat
    rstDVD.Update
  End If
End Property


Public Property Get bytDiscFormat() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bytDiscFormat
    bytDiscFormat = mvarbytDiscFormat
End Property



Public Property Let blnDVDRom(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnDVDRom = 5
    mvarblnDVDRom = vData
    rstDVD![blnDVDRom] = vData
    rstDVD.Update
End Property


Public Property Get blnDVDRom() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnDVDRom
    blnDVDRom = mvarblnDVDRom
End Property



Public Property Let blnStarBios(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnStarBios = 5
    mvarblnStarBios = vData
    rstDVD![blnStarBios] = vData
    rstDVD.Update
End Property


Public Property Get blnStarBios() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnStarBios
        blnStarBios = mvarblnStarBios
End Property


Public Property Let blnTheatricalTrailer(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnTheatricalTrailer = 5
    mvarblnTheatricalTrailer = vData
    rstDVD![blnTheatricalTrailer] = vData
    rstDVD.Update
End Property

Public Property Get blnTheatricalTrailer() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnTheatricalTrailer
        blnTheatricalTrailer = mvarblnTheatricalTrailer
End Property



Public Property Let blnDeletedScenes(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnDeletedScenes = 5
    mvarblnDeletedScenes = vData
    rstDVD![blnDeletedScenes] = vData
    rstDVD.Update
End Property


Public Property Get blnDeletedScenes() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnDeletedScenes
    blnDeletedScenes = mvarblnDeletedScenes
End Property



Public Property Let blnCommentary(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnCommentary = 5
    mvarblnCommentary = vData
    rstDVD![blnCommentary] = vData
    rstDVD.Update
End Property


Public Property Get blnCommentary() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnCommentary
    blnCommentary = mvarblnCommentary
End Property

Public Property Let blnMakingOf(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnMakingOf = 5
    mvarblnMakingOf = vData
    rstDVD![blnMakingOf] = vData
    rstDVD.Update
End Property


Public Property Get blnMakingOf() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnMakingOf
    blnMakingOf = mvarblnMakingOf
End Property



Public Property Let blnAnimatedMenus(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnAnimatedMenus = 5
    mvarblnAnimatedMenus = vData
    rstDVD![blnAnimatedMenus] = vData
    rstDVD.Update
End Property


Public Property Get blnAnimatedMenus() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnAnimatedMenus
    blnAnimatedMenus = mvarblnAnimatedMenus
End Property



Public Property Let blnSceneAccess(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnSceneAccess = 5
    mvarblnSceneAccess = vData
    rstDVD![blnSceneAccess] = vData
    rstDVD.Update
End Property

Public Property Get blnSceneAccess() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnSceneAccess
        blnSceneAccess = mvarblnSceneAccess
End Property



Public Property Let blnAudioOther(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnAudioOther = 5
    mvarblnAudioOther = vData
    rstDVD![blnAudioOther] = vData
    rstDVD.Update
End Property


Public Property Get blnAudioOther() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnAudioOther
    blnAudioOther = mvarblnAudioOther
End Property



Public Property Let blnSDDS(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnSDDS = 5
    mvarblnSDDS = vData
    rstDVD![blnSDDS] = vData
    rstDVD.Update
End Property

Public Property Get blnSDDS() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnSDDS
        blnSDDS = mvarblnSDDS
End Property



Public Property Let blnDTS(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnDTS = 5
    mvarblnDTS = vData
    rstDVD![blnDTS] = vData
    rstDVD.Update
End Property


Public Property Get blnDTS() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnDTS
    blnDTS = mvarblnDTS
End Property



Public Property Let blnDDEx(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnDDEx = 5
    mvarblnDDEx = vData
    rstDVD![blnDDEx] = vData
    rstDVD.Update
End Property


Public Property Get blnDDEx() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnDDEx
    blnDDEx = mvarblnDDEx
End Property



Public Property Let blnDD51(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnDD51 = 5
    mvarblnDD51 = vData
    rstDVD![blnDD51] = vData
    rstDVD.Update
End Property


Public Property Get blnDD51() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnDD51
    blnDD51 = mvarblnDD51
End Property



Public Property Let blnDolbyProLogic(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnDolbyProLogic = 5
    mvarblnDolbyProLogic = vData
    rstDVD![blnDolbyProLogic] = vData
    rstDVD.Update
End Property


Public Property Get blnDolbyProLogic() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnDolbyProLogic
    blnDolbyProLogic = mvarblnDolbyProLogic
End Property



Public Property Let blnDolbySurround(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnDolbySurround = 5
    mvarblnDolbySurround = vData
    rstDVD![blnDolbySurround] = vData
    rstDVD.Update
End Property


Public Property Get blnDolbySurround() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnDolbySurround
    blnDolbySurround = mvarblnDolbySurround
End Property



Public Property Let blnStereo(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnStereo = 5
    mvarblnStereo = vData
    rstDVD![blnStereo] = vData
    rstDVD.Update
End Property

Public Property Get blnStereo() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnStereo
   blnStereo = mvarblnStereo
End Property

Public Property Let blnSubtitleOther(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnSubtitleOther = 5
    mvarblnSubtitleOther = vData
    rstDVD![blnSubtitleOther] = vData
    rstDVD.Update
End Property

Public Property Get blnSubtitleOther() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnSubtitleOther
   blnSubtitleOther = mvarblnSubtitleOther
End Property

Public Property Let blnChinese(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnChinese = 5
    mvarblnChinese = vData
    rstDVD![blnChinese] = vData
    rstDVD.Update
End Property


Public Property Get blnChinese() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnChinese
    blnChinese = mvarblnChinese
End Property



Public Property Let blnJapanese(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnJapanese = 5
    mvarblnJapanese = vData
    rstDVD![blnJapanese] = vData
    rstDVD.Update
End Property


Public Property Get blnJapanese() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnJapanese
    blnJapanese = mvarblnJapanese
End Property



Public Property Let blnPortugese(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnPortugese = 5
    mvarblnPortugese = vData
    rstDVD![blnPortugese] = vData
    rstDVD.Update
End Property


Public Property Get blnPortugese() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnPortugese
    blnPortugese = mvarblnPortugese
End Property



Public Property Let blnSpanish(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnSpanish = 5
    mvarblnSpanish = vData
    rstDVD![blnSpanish] = vData
    rstDVD.Update
End Property


Public Property Get blnSpanish() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnSpanish
    blnSpanish = mvarblnSpanish
End Property



Public Property Let blnGerman(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnGerman = 5
    mvarblnGerman = vData
    rstDVD![blnGerman] = vData
    rstDVD.Update
End Property


Public Property Get blnGerman() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnGerman
    blnGerman = mvarblnGerman
End Property



Public Property Let blnFrench(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnFrench = 5
    mvarblnFrench = vData
    rstDVD![blnFrench] = vData
    rstDVD.Update
End Property


Public Property Get blnFrench() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnFrench
    blnFrench = mvarblnFrench
End Property



Public Property Let blnEnglish(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnEnglish = 5
    mvarblnEnglish = vData
    rstDVD![blnEnglish] = vData
    rstDVD.Update
End Property


Public Property Get blnEnglish() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnEnglish
    blnEnglish = mvarblnEnglish
End Property



Public Property Let blnNTSCPAL(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnNTSCPAL = 5
  If rstDVD.State = adStateOpen Then
    mvarblnNTSCPAL = vData
    rstDVD![blnNTSCPAL] = vData
    rstDVD.Update
  End If
End Property


Public Property Get blnNTSCPAL() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnNTSCPAL
    blnNTSCPAL = mvarblnNTSCPAL
End Property



Public Property Let dblRatio(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dblRatio = 5
    mvardblRatio = vData
    rstDVD![dblRatio] = vData
    rstDVD.Update
End Property


Public Property Get dblRatio() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dblRatio
    dblRatio = mvardblRatio
End Property



Public Property Let bln169(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bln169 = 5
    mvarbln169 = vData
    rstDVD![bln169] = vData
    rstDVD.Update
End Property


Public Property Get bln169() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bln169
    bln169 = mvarbln169
End Property



Public Property Let blnPanScan(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnPanScan = 5
    mvarblnPanScan = vData
    rstDVD![blnPanScan] = vData
    rstDVD.Update
End Property


Public Property Get blnPanScan() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnPanScan
    blnPanScan = mvarblnPanScan
End Property



Public Property Let blnFullFrame(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnFullFrame = 5
    mvarblnFullFrame = vData
    rstDVD![blnFullFrame] = vData
    rstDVD.Update
End Property


Public Property Get blnFullFrame() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnFullFrame
    blnFullFrame = mvarblnFullFrame
End Property



Public Property Let blnWidescreen(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.blnWidescreen = 5
    mvarblnWidescreen = vData
    rstDVD![blnWidescreen] = vData
    rstDVD.Update
End Property

Public Property Get blnWidescreen() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.blnWidescreen
    blnWidescreen = mvarblnWidescreen
End Property



Public Property Let strDirector(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strDirector = 5
    mvarstrDirector = vData
    rstDVD![strDirector] = vData
    rstDVD.Update
End Property


Public Property Get strDirector() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strDirector
    strDirector = mvarstrDirector
End Property



Public Property Let strStudio(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strStudio = 5
    mvarstrStudio = vData
    rstDVD![strStudio] = vData
    rstDVD.Update
End Property


Public Property Get strStudio() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strStudio
    strStudio = mvarstrStudio
End Property



Public Property Let curCost(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.curCost = 5
    mvarcurCost = vData
    rstDVD![curCost] = vData
    rstDVD.Update
End Property


Public Property Get curCost() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.curCost
    curCost = mvarcurCost
End Property



Public Property Let datDatePurchased(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.datDatePurchased = 5
    mvardatDatePurchased = vData
    rstDVD![datDatePurchased] = vData
    rstDVD.Update
End Property


Public Property Get datDatePurchased() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.datDatePurchased
    datDatePurchased = mvardatDatePurchased
End Property



Public Property Let strLocationPurchased(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strLocationPurchased = 5
    mvarstrLocationPurchased = vData
    rstDVD![strLocationPurchased] = vData
    rstDVD.Update
End Property


Public Property Get strLocationPurchased() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strLocationPurchased
    strLocationPurchased = mvarstrLocationPurchased
End Property



Public Property Let lngCurrentLocation(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngCurrentLocation = 5
  If mvarlngCurrentLocation <> vData Then
    mvarlngCurrentLocation = vData
    rstDVD![lngCurrentLocation] = vData
    rstDVD.Update
  End If
End Property


Public Property Get lngCurrentLocation() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngCurrentLocation
    lngCurrentLocation = mvarlngCurrentLocation
End Property



Public Property Let lngCaseType(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngCaseType = 5
  If mvarlngCaseType <> vData Then
    mvarlngCaseType = vData
    rstDVD![lngCaseType] = vData
    rstDVD.Update
  End If
End Property


Public Property Get lngCaseType() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngCaseType
    lngCaseType = mvarlngCaseType
End Property



Public Property Let lngRating(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngRating = 5
  If mvarlngRating <> vData Then
    mvarlngRating = vData
    rstDVD![lngRating] = vData
    rstDVD.Update
  End If
End Property


Public Property Get lngRating() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngRating
    lngRating = mvarlngRating
End Property



Public Property Let lngRegion(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngRegion = 5
  If mvarlngRegion <> vData Then
    mvarlngRegion = vData
    rstDVD![lngRegion] = vData
    rstDVD.Update
  End If
End Property


Public Property Get lngRegion() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngRegion
    lngRegion = mvarlngRegion
End Property



Public Property Let datDVDRelease(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.datDVDRelease = 5
  If mvardatDVDRelease <> vData Then
    mvardatDVDRelease = vData
    rstDVD![datDVDRelease] = vData
    rstDVD.Update
  End If
End Property


Public Property Get datDVDRelease() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.datDVDRelease
    datDVDRelease = mvardatDVDRelease
End Property



Public Property Let datMovieYear(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.datMovieYear = 5
    mvardatMovieYear = vData
    rstDVD![datMovieYear] = vData
    rstDVD.Update
End Property


Public Property Get datMovieYear() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.datMovieYear
    datMovieYear = mvardatMovieYear
End Property



Public Property Let bytUserReview(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bytUserReview = 5
    mvarbytUserReview = vData
    rstDVD![bytUserReview] = vData
    rstDVD.Update
End Property


Public Property Get bytUserReview() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bytUserReview
    bytUserReview = mvarbytUserReview
End Property



Public Property Let lngGenre(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngGenre = 5
  If mvarlngGenre <> vData Then
    mvarlngGenre = vData
    rstDVD![lngGenre] = vData
    rstDVD.Update
  End If

End Property


Public Property Get lngGenre() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngGenre
    lngGenre = mvarlngGenre
End Property

Public Property Let strTitle(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strTitle = 5
    mvarstrTitle = vData
    rstDVD![strTitle] = vData
    rstDVD.Update
End Property

Public Property Get strTitle() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strTitle
    strTitle = mvarstrTitle
End Property

Public Property Let lngID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngID = 5
    mvarlngID = vData
    rstDVD![lngID] = vData
    rstDVD.Update
End Property


Public Property Get lngID() As Long
Attribute lngID.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngID
    lngID = mvarlngID
End Property

Public Property Let intRunningTime(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngID = 5
    mvarintRunningTime = vData
    rstDVD![intRunningTime] = vData
    rstDVD.Update
End Property


Public Property Get intRunningTime() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngID
    intRunningTime = mvarintRunningTime
End Property

Private Function returnByte(blnTmp As Boolean) As Byte
  If blnTmp = True Then
    returnByte = 1
  Else
    returnByte = 0
  End If
End Function

