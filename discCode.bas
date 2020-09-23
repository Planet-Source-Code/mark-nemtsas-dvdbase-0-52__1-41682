Attribute VB_Name = "discCode"
Option Explicit


Function updateTreeView(strSortKey As String, cboTmp As ComboBox)
  Dim nNode As Node, nParentNode As Node
  Dim strName As String, strText As String, strNewParent As String
  Select Case strSortKey
    Case "Genre"
      If frmMain.optGenreSort.Value <> True Then
        Exit Function
      End If
    Case "Region"
      If frmMain.optRegionSort.Value <> True Then
        Exit Function
      End If
    Case "Rating"
      If frmMain.optRatingSort.Value <> True Then
        Exit Function
      End If
  End Select
  Set nNode = frmMain.treDiscs.SelectedItem
  If nNode.Parent.Key = cboTmp.List(cboTmp.ListIndex) Then Exit Function 'only do something if something has changed
  strName = nNode.Key
  strText = nNode.Text
  Set nParentNode = nNode.Parent
  If nNode.Parent.Children = 1 Then nNode.Parent.Image = "Closed"
  Set nNode = Nothing
  frmMain.treDiscs.Nodes.Remove strName
  If cboTmp.ListIndex <> -1 Then
    strNewParent = cboTmp.List(cboTmp.ListIndex)
  Else
    strNewParent = "No " & strSortKey
  End If
  Set nNode = frmMain.treDiscs.Nodes.Add(strNewParent, tvwChild, strName, strText, "Disc")
  Set frmMain.treDiscs.SelectedItem = nNode
  nParentNode.Expanded = True
End Function

Function updateTreeViewDragDrop(strSortKey As String, cboTmp As ComboBox, nDragNode As Node, nDropNode As Node)
  Dim nNode As Node
  Dim strName As String, strText As String, strNewParent As String
  Select Case strSortKey
    Case "Genre"
      If frmMain.optGenreSort.Value <> True Then
        Exit Function
      End If
    Case "Region"
      If frmMain.optRegionSort.Value <> True Then
        Exit Function
      End If
    Case "Rating"
      If frmMain.optRatingSort.Value <> True Then
        Exit Function
      End If
  End Select
  If nDragNode.Parent.Key = nDropNode.Key Then Exit Function  'only do something if something has changed
  cboTmp = nDropNode.Key
End Function

Function enableDiscDisplay()
  Dim ctrTmp As Control
  
  For Each ctrTmp In frmMain.Controls
    If ctrTmp.Name <> "ilsTreeView" Then ctrTmp.Enabled = True
  Next ctrTmp
End Function

Function enableAddDiscDisplay()
  Dim ctrTmp As Control
  
  For Each ctrTmp In frmMain.Controls
    Select Case ctrTmp.Name
      Case "treDiscs"
        ctrTmp.Enabled = True
      Case "ilsTreeView"
      Case "btnDelete"
        ctrTmp.Enabled = True
      Case "btnNew"
        ctrTmp.Enabled = True
      Case "txtTitle"
        ctrTmp.Enabled = True
      Case "tabMain"
        ctrTmp.Enabled = True
      Case "fraGeneral"
        ctrTmp.Enabled = True
      Case "fraSort"
        ctrTmp.Enabled = True
      Case "lblTitle"
        ctrTmp.Enabled = True
      Case "mnuAbout"
        ctrTmp.Enabled = True
      Case Else
        ctrTmp.Enabled = False
    End Select
  Next ctrTmp
End Function

Function disableDiscDisplay()
  Dim ctrTmp As Control
  
  For Each ctrTmp In frmMain.Controls
    Select Case ctrTmp.Name
      Case "treDiscs"
        ctrTmp.Enabled = True
      Case "ilsTreeView"
      Case "btnDelete"
        ctrTmp.Enabled = True
      Case "btnNew"
        ctrTmp.Enabled = True
      Case "tabMain"
        ctrTmp.Enabled = True
      Case "mnuAbout"
        ctrTmp.Enabled = True
      Case "fraSort"
        ctrTmp.Enabled = True
      Case "optNoSort"
        ctrTmp.Enabled = True
      Case "optGenreSort"
        ctrTmp.Enabled = True
      Case "optRegionSort"
        ctrTmp.Enabled = True
      Case "optRatingSort"
        ctrTmp.Enabled = True
      Case Else
        ctrTmp.Enabled = False
    End Select
  Next ctrTmp
End Function

Function resetAllFields()
  Dim intLoop As Integer
  
  frmMain.txtTitle = ""
  frmMain.cboGenre.ListIndex = -1
  For intLoop = 0 To 4
    frmMain.picStar(intLoop).Visible = False
  Next intLoop
  frmMain.txtMovieYear = ""
  frmMain.txtDVDRelease = ""
  frmMain.cboRegion.ListIndex = -1
  frmMain.cboRating.ListIndex = -1
  frmMain.cboCaseType.ListIndex = -1
  frmMain.cboCurrentLocation.ListIndex = -1
  frmMain.txtLocationPurchased = ""
  frmMain.txtDatePurchased = ""
  frmMain.txtCost = ""
  frmMain.txtStudio = ""
  frmMain.txtDirector = ""
  frmMain.chkWidescreen = False
  frmMain.chkFullFrame = False
  frmMain.chkPanScan = False
  frmMain.chk169 = False
  frmMain.txtRatio = ""
  frmMain.txtRunningTime = ""
  frmMain.optNTSC = True
  frmMain.optPAL = False
  frmMain.chkEnglish = False
  frmMain.chkFrench = False
  frmMain.chkGerman = False
  frmMain.chkSpanish = False
  frmMain.chkPortugese = False
  frmMain.chkJapanese = False
  frmMain.chkChinese = False
  frmMain.chkSubTitleOther = False
  frmMain.chkStereo = False
  frmMain.chkDolbySurround = False
  frmMain.chkDolbyProLogic = False
  frmMain.chkdd51 = False
  frmMain.chkDDEx = False
  frmMain.chkDTS = False
  frmMain.chkSDDS = False
  frmMain.chkAudioOther = False
  frmMain.chkSceneAccess = False
  frmMain.chkAnimatedMenus = False
  frmMain.chkMakingOf = False
  frmMain.chkCommentary = False
  frmMain.chkDeletedScenes = False
  frmMain.chkTheatricalTrailer = False
  frmMain.chkBios = False
  frmMain.optDualLayer = True
  frmMain.optDualSided = False
  frmMain.optFlipper = False
End Function

Public Function getLatestDVD() As Long
  Dim rstDVDs As New ADODB.Recordset
  
  Set rstDVDs = returnRS(cmdSelectLatestDVD)
  If rstDVDs.EOF <> True Then
    rstDVDs.MoveLast
    getLatestDVD = rstDVDs![lngID]
  Else
    MsgBox "Error: Your database seems to be corrupted, no records found for clsDVD.setLatestDVD"
  End If
  rstDVDs.Close
  Set rstDVDs = Nothing
End Function

Function returnComboLocation(lngArray() As Long, intArraySize As Integer, lngID As Long) As Integer
  Dim intLoop As Integer
  returnComboLocation = -1
  For intLoop = 1 To intArraySize
    If lngArray(intLoop) = lngID Then
      returnComboLocation = intLoop
      Exit Function
    End If
  Next intLoop
End Function

Function parseTime(strTime As String) As Integer
  Dim intPosition As Integer
  Dim strHours As String, strMinutes As String, strSeconds As String
  Dim intHours As Integer, intMinutes As Integer, intSeconds As Integer
  On Error GoTo errorHandler
  If Len(strTime) = 0 Then
    parseTime = -1
    Exit Function
  End If
  If IsNumeric(strTime) = True Then
    If CLng(strTime) * 3600 > 65000 Then
      parseTime = -1
      Exit Function
    End If
    intHours = CInt(strTime)
    parseTime = intHours * 3600
    Exit Function
  End If
  intPosition = InStr(1, strTime, ":")
  If intPosition = 0 Then
    parseTime = -1
    Exit Function
  End If
  If intPosition > 1 Then
    strHours = Left(strTime, intPosition - 1)
    strTime = Right(strTime, Len(strTime) - intPosition)
    If IsNumeric(strHours) = True Then
      intHours = CInt(strHours) * 3600
    Else
      parseTime = -1
      Exit Function
    End If
  End If
  If IsNumeric(strTime) = True Then
    If CLng(strTime) * 60 + intHours > (65000) Then
      parseTime = -1
      Exit Function
    End If
    intMinutes = CInt(strTime) * 60
    parseTime = intHours + intMinutes
    Exit Function
  End If
  intPosition = InStr(1, strTime, ":")
  If intPosition = 0 Then
    parseTime = -1
    Exit Function
  End If
  If intPosition > 1 Then
    strMinutes = Left(strTime, intPosition - 1)
    strTime = Right(strTime, Len(strTime) - intPosition)
    If IsNumeric(strMinutes) = True Then
      intMinutes = CInt(strMinutes) * 60
    Else
      parseTime = -1
      Exit Function
    End If
  End If
  If IsNumeric(strTime) = True Then
    If CLng(strTime) + intHours + intMinutes > (65000) Then
      parseTime = -1
      Exit Function
    End If
    intSeconds = CInt(strTime)
    parseTime = intHours + intMinutes + intSeconds
    Exit Function
  Else
    parseTime = -1
    Exit Function
  End If
  
  Exit Function
errorHandler:
  MsgBox "Congratulations you broke #discCode#parseTime with error: " & Err.Number & " " & Err.Description & Chr(10) & _
  "Send me a tutorial on writing recursive functions and I'll make a more robust parser"
  Err.Clear
  parseTime = -1
  Exit Function
End Function
