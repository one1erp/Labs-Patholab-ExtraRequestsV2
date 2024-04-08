VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.UserControl ExtraRequestsCtrlV2 
   ClientHeight    =   9570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17115
   ScaleHeight     =   9570
   ScaleWidth      =   17115
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   480
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   14160
      Top             =   480
   End
   Begin VB.Frame Frame2 
      Caption         =   "Report Removal from Archive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   7080
      TabIndex        =   5
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdReport 
         Caption         =   "Report Selected"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtEntityBarcode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblMicrotom 
         Caption         =   "Microtome: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Barcode Entity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Presentation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdShowRequests 
         Caption         =   "Show    Requests"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ListBox lstEntityTypes 
         Height          =   690
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   2652
      End
      Begin VB.Label Label1 
         Caption         =   "Entity Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   12726
      _Version        =   393216
   End
End
Attribute VB_Name = "ExtraRequestsCtrlV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements LSExtensionWindowLib.IExtensionWindow
Implements LSExtensionWindowLib.IExtensionWindow2

Option Explicit


Private ProcessXML As LSSERVICEPROVIDERLib.NautilusProcessXML
Private NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
Private NtlsSite As LSExtensionWindowLib.IExtensionWindowSite2
Private NtlsUser As LSSERVICEPROVIDERLib.NautilusUser

'Ashi - Sorting
Dim m_SortColumn As Integer

Dim m_SortOrder As Integer

Private con As New ADODB.Connection
Private entityType As String

'holds the entities of the presented list
'key  - the entity name
'item - a collection of locations on the grid for this entity
Private dicEntities As New Dictionary

'holds the extra_request_data_id against the entity name:
Private dicBarcodeEntities As New Dictionary

Private Const MARK_SELECTED = &HC0FFFF

'used for the printing of the
'grid to the printer:
Private RowSize As Long
Private ColSize As Long
Private LeftOffset As Long
Private TopOffset As Long

Private Const MAX_DIGITS_PER_CELL = 24
Public RunFromWindow As Boolean
Private Const HI = "היסטוכימיה"
Private Const IM = "אימונוהיסטוכימיה"

Public Event CloseClicked()
'data of the font to be used when printing the aliquot names.
'to be read from a relevant phrase
Private strFontName As String
Private iFontSize As Integer
Private isFontBold As Boolean

Private Const MAX_LINES = 20

Private dicOperatorAllowedToReportReserveSlides As New Dictionary

Private sdg_log As New SdgLog.CreateLog


Private Sub cmdClose_Click()

13060     If RunFromWindow Then
13070         RaiseEvent CloseClicked
13080     Else
13090         NtlsSite.CloseWindow
13100     End If
End Sub

'print the grid contents;
'on each page print up to MAX_LINES rows of the grid;
'the column headers are printed on each page;
Private Sub cmdPrint_Click()
10    On Error GoTo ERR_cmdPrint_Click

          Dim I As Integer
          Dim iFirstLine As Integer
          Dim iLastLine As Integer
          
20        I = 1
          
30        While I < grid.Rows
40            iFirstLine = I
50            iLastLine = IIf(I + MAX_LINES >= grid.Rows, grid.Rows - 1, I + MAX_LINES - 1)

60            Call PrintToPrinter(iFirstLine, iLastLine, MAX_LINES)

70            I = I + MAX_LINES
80        Wend
          
90        lblMicrotom.Visible = False
  
          

100       Exit Sub
ERR_cmdPrint_Click:
110   MsgBox "ERR_cmdPrint_Click" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


'barcode all presented entities:
Private Sub BarcodeAll()
10    On Error GoTo ERR_BarcodeAll

          Dim I As Integer

20        For I = 0 To dicEntities.Count - 1
30            Call BarcodeEntity(CStr(dicEntities.Keys(I)))
40        Next I

50        Exit Sub
ERR_BarcodeAll:
60    MsgBox "ERR_BarcodeAll" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub cmdReport_Click()
10    On Error GoTo ERR_cmdReport_Click
          Dim I As Integer
          Dim k As Integer
          Dim sql As String
          Dim rs As Recordset
          Dim strLastEntity As String
          Dim strExternalRef As String
          Dim strSdgId As String
         
20        For I = 0 To dicBarcodeEntities.Count - 1
30            sql = " update lims_sys.u_extra_request_data_user"
40            sql = sql & " set u_status = 'P'"
50            sql = sql & " where u_extra_request_data_id = " & dicBarcodeEntities.Keys(I)
60            con.Execute (sql)
              'update sdg_log table (once per entity)
70            If strLastEntity <> CStr(dicBarcodeEntities.Items(I)) Then

80                strSdgId = GetSdgForEntity(CStr(dicBarcodeEntities.Items(I)), entityType)

      '            'sql = " select ru.U_SDG_ID"
      '            sql = " select ru.U_EXTERNAL_REFERENCE "
      '            sql = sql & " from lims_sys.u_extra_request_user ru,"
      '            sql = sql & "      lims_sys.u_extra_request_data_user rdu"
      '            sql = sql & " where ru.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
      '            sql = sql & " and   rdu.U_EXTRA_REQUEST_DATA_ID=" & dicBarcodeEntities.Keys(i)
      '            Set rs = con.Execute(sql)
      '
      '            strExternalRef = nte(rs("U_EXTERNAL_REFERENCE"))
      '
      '            'get the right SDG
      '            '(considering this might be a revision):
      '            sql = " select d.SDG_ID, d.name"
      '            sql = sql & " from lims_sys.sdg d"
      '            sql = sql & " where d.EXTERNAL_REFERENCE='" & strExternalRef & "'"
      '            sql = sql & " and   instr(d.name, 'V')=0"
      '            sql = sql & " order by d.sdg_id desc"
      '            Set rs = con.Execute(sql)
90                Call sdg_log.InsertLog(CLng(strSdgId), _
                                         "EXTRA.STORAGE", _
                                         CStr(dicBarcodeEntities.Items(I)))
100           End If
110           strLastEntity = dicBarcodeEntities.Items(I)
              
120           Call UpdateArchive(entityType, CStr(dicBarcodeEntities.Items(I)), "F")
130           k = k + 1
140       Next I
150       If k = 1 Then
160           MsgBox "one record was updated"
170       Else
180           MsgBox CStr(k) & " records were updated"
190       End If
          
200       dicBarcodeEntities.RemoveAll
210       cmdReport.Enabled = False
220       Call cmdShowRequests_Click
230       If dicEntities.Count = 0 Then
            '  Call lstEntityTypes.RemoveItem(lstEntityTypes.ListIndex)
240       End If
          
250       If lstEntityTypes.ListCount = 0 Then
260           cmdShowRequests.Enabled = False
270       End If
280       lblMicrotom.Visible = False
          
290       Exit Sub
ERR_cmdReport_Click:
300   MsgBox "ERR_cmdReport_Click" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub



Private Function GetSdgForEntity(strEntityName As String, strEntityType As String) As String
10    On Error GoTo ERR_GetSdgForEntity

          Dim rs As Recordset
          Dim sql As String


          
20            sql = " select s.SDG_ID"
30            sql = sql & " from lims_sys.sample s,"
40            sql = sql & "      lims_sys.aliquot a"
50            sql = sql & " where a.SAMPLE_ID=s.SAMPLE_ID"
60            sql = sql & " and   a.NAME='" & strEntityName & "'"
          

70        Set rs = con.Execute(sql)
          
80        If Not rs.EOF Then
90            GetSdgForEntity = nte(rs("SDG_ID"))
100       End If

110       Exit Function
ERR_GetSdgForEntity:
120   MsgBox "ERR_GetSdgForEntity" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function


'clear both dictionaries
'present the selection
'build the dictionary of selected items: ENTITY_NAME -> ROWS_IN_THE_GRID
Private Sub cmdShowRequests_Click()
10    On Error GoTo ERR_cmdShowRequests_Click
          Dim sql As String
          Dim rs As Recordset
          Dim iRows As Integer
          Dim I As Integer
          Dim s As String
          Dim strColorGroup As String
          Dim dicLocations As Dictionary
          
       
20        entityType = lstEntityTypes.Text
30        If entityType = "LBC" Then entityType = "Sample"
          
              
40        Call dicEntities.RemoveAll
50        Call dicBarcodeEntities.RemoveAll
60        Call InitializeGrid
          
   Dim partType As String
70       If entityType = HI Then
80        partType = "H','O"
90        Else
100       partType = "I"
110       End If
     'CHECK H OR I PATHOLAB
          
120          sql = "select rd.U_EXTRA_REQUEST_DATA_ID ID, nvl(du.U_PATHOLAB_NUMBER,'') ||substr(rd.name,11) as PATHOLAB_NAME , rd.NAME ENTITY_NAME,"
130          sql = sql & "  r.NAME ACTION, rdu.U_REQUEST_DETAILS,"
140          sql = sql & "  o.Name , ru.U_CREATED_ON"
150          sql = sql & "  from lims_sys.u_extra_request_data rd,"
160          sql = sql & "  lims_sys.u_extra_request_data_user rdu,"
170             sql = sql & "  lims_sys.u_parts_user p,"
180          sql = sql & "  lims_sys.u_extra_request r,"
190          sql = sql & "  lims_sys.u_extra_request_user ru,"
200          sql = sql & "  lims_sys.operator o,"
210          sql = sql & "  lims_sys.aliquot  al ,"
220          sql = sql & "  lims_sys.aliquot_user alu, "
230          sql = sql & "  lims_sys.sdg_user du,lims_sys.aliquot  alc ,lims_sys.aliquot_user alcu "
240          sql = sql & "  Where rd.U_EXTRA_REQUEST_DATA_ID = rdu.U_EXTRA_REQUEST_DATA_ID"
250          sql = sql & "  and r.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
260          sql = sql & "  and r.U_EXTRA_REQUEST_ID=ru.U_EXTRA_REQUEST_ID"
270          sql = sql & "  and o.OPERATOR_ID=ru.U_CREATED_BY"
280         sql = sql & "  and rdu.U_ENTITY_TYPE='Block'"
290          sql = sql & "  and rdu.U_STATUS='V'"
300          sql = sql & "  and du.SDG_ID=ru.U_SDG_ID"
310          sql = sql & "  AND al.ALIQUOT_ID=alu.ALIQUOT_ID"
320          sql = sql & " AND alc.ALIQUOT_ID=alcu.ALIQUOT_ID and alc.name=RDU.U_SLIDE_NAME "
330          sql = sql & "  AND alcu.U_COLOR_TYPE=P.U_STAIN and P.U_PART_TYPE  IN ('" & partType & "')"
340          sql = sql & "  AND alu.U_ALIQUOT_STATION <'5'"
350          sql = sql & "  AND al.NAME = SUBSTR(rd.NAME,1,INSTR(rd.NAME,';',1)-1)" 'join to specified aliquot
360          sql = sql & "  order by rd.NAME, rd.U_EXTRA_REQUEST_DATA_ID"
370          Set rs = con.Execute(sql)
          'end
        '  MsgBox sql
380
390       iRows = 1
          
400       While Not rs.EOF
410           iRows = iRows + 1
          
420           grid.Rows = iRows
430           grid.col = 0
440           grid.row = grid.Rows - 1
              
450           For I = 0 To rs.Fields.Count - 1
460               grid.col = I
470               grid.CellAlignment = vbLeftJustify
                  
480               s = Trim(CleanSemicolon(nte(rs.Fields(I).Value)))
                 
490              If nte(rs.Fields(I).Name) = "U_REQUEST_DETAILS" Then
500                   strColorGroup = GetColorGroup(s)
510               Else
520                   strColorGroup = ""
530               End If
                  
540               If strColorGroup <> "" Then
550                   s = s & " (" & strColorGroup & ")"
560               End If
                  
570               grid.Text = s
      '            grid.Text = Trim(CleanSemicolon(nte(rs.Fields(i).Value)))
580           Next I
                         
590           s = Trim(CleanSemicolon(nte(rs("ENTITY_NAME"))))
                         
600           If dicEntities.Exists(s) Then
610               Set dicLocations = dicEntities(s)
620               Call dicLocations.Add(CStr(grid.row), "")
              
             '     dicEntities(s) = dicEntities(s) & "," & CStr(grid.row)
630           Else
640               Set dicLocations = New Dictionary
650               Call dicLocations.Add(CStr(grid.row), "")
660               Call dicEntities.Add(s, dicLocations)
      '            Call dicEntities.Add(s, CStr(grid.row))
670           End If
                         
                         
680           If ExistRemark(grid.TextMatrix(grid.row, 0)) Then
690               grid.col = 0
700               grid.CellFontBold = True
710           End If
                         
720           rs.MoveNext
730       Wend
          
740       If dicEntities.Count > 0 Then
              'Call InitReportFrame(True)
750           txtEntityBarcode.Enabled = True
760           cmdPrint.Enabled = True
770       End If
          
         ' DEBUG_SHOW_DICTIONARY
          
780       lblMicrotom.Visible = False
790       cmdReport.Enabled = False
          
          'for reserve slides - all should be auto-barcoded:
800       If entityType = "Reserve-Slide" Then
          'And dicOperatorAllowedToReportReserveSlides.Exists(CStr(NtlsUser.GetOperatorId)) Then
810           Call BarcodeAll
820       End If
          
830       Exit Sub
ERR_cmdShowRequests_Click:
840   MsgBox "ERR_cmdShowRequests_Click" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub





Public Function IExtensionWindow_CloseQuery() As Boolean
          'Happens when the user close the window
          'Call UnloadRequest
         Set con = Nothing 'ASHI 18-04-21
10        IExtensionWindow_CloseQuery = True
End Function

Public Function IExtensionWindow_DataChange() As LSExtensionWindowLib.WindowRefreshType
10        IExtensionWindow_DataChange = windowRefreshNone
End Function

Public Function IExtensionWindow_GetButtons() As LSExtensionWindowLib.WindowButtonsType
10        IExtensionWindow_GetButtons = windowButtonsNone
End Function

Public Sub IExtensionWindow_Internationalise()
End Sub

Public Sub IExtensionWindow_PreDisplay()

1690      Set con = New ADODB.Connection
Dim cs As String
          cs = NtlsCon.GetADOConnectionString

       If NtlsCon.GetServerIsProxy Then
          cs = "Provider=OraOLEDB.Oracle;Data Source=" & _
            NtlsCon.GetServerDetails & ";User id=/;Persist Security Info=True;"
        End If

        


60        con.Open cs
70        con.CursorLocation = adUseClient
80        con.Execute "SET ROLE LIMS_USER"

90        Call ConnectSameSession(CDbl(NtlsCon.GetSessionId))
          
100       Set sdg_log.con = con
110       sdg_log.Session = CDbl(NtlsCon.GetSessionId)
End Sub

Public Sub IExtensionWindow_refresh()
    'Code for refreshing the window
End Sub

Public Sub IExtensionWindow_RestoreSettings(ByVal hKey As Long)
End Sub

Public Function IExtensionWindow_SaveData() As Boolean
End Function

Public Sub IExtensionWindow_SaveSettings(ByVal hKey As Long)
End Sub

Public Sub IExtensionWindow_SetParameters(ByVal parameters As String)
10    On Error GoTo ERR_IExtensionWindow_SetParameters

          Dim strMain As String
          Dim s As String
          
20        strMain = parameters
          
30        While strMain <> ""
40            s = getNextStr(strMain, " ")
50            Call dicOperatorAllowedToReportReserveSlides.Add(s, "")
60        Wend
          
70        Exit Sub
ERR_IExtensionWindow_SetParameters:
80    MsgBox "ERR_IExtensionWindow_SetParameters" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Public Sub IExtensionWindow_SetServiceProvider(ByVal serviceProvider As Object)
          Dim sp As LSSERVICEPROVIDERLib.NautilusServiceProvider
10        Set sp = serviceProvider
20        Set ProcessXML = sp.QueryServiceProvider("ProcessXML")
30        Set NtlsCon = sp.QueryServiceProvider("DBConnection")
40        Set NtlsUser = sp.QueryServiceProvider("User")
End Sub

Public Sub IExtensionWindow_SetSite(ByVal Site As Object)
10        Set NtlsSite = Site
20        NtlsSite.SetWindowInternalName ("ExtraRequsts")
30        NtlsSite.SetWindowRegistryName ("ExtraRequsts")
40        Call NtlsSite.SetWindowTitle("Pathology Extra Requsts")
End Sub

Public Sub IExtensionWindow_Setup()
10    On Error GoTo ERR_IExtensionWindow_Setup
          Dim rs As Recordset
          Dim sql As String
    
20        cmdReport.Enabled = False
30        cmdPrint.Enabled = False
40        txtEntityBarcode.Enabled = False
50        Call InitializeGrid
60        Call InitEntityTypes
          
          
70        Exit Sub
ERR_IExtensionWindow_Setup:
80    MsgBox "ERR_IExtensionWindow_Setup" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Public Function IExtensionWindow_ViewRefresh() As LSExtensionWindowLib.WindowRefreshType
10        IExtensionWindow_ViewRefresh = windowRefreshNone
End Function

Private Sub ConnectSameSession(ByVal aSessionID)
          Dim aProc As New ADODB.Command
          Dim aSession As New ADODB.Parameter
          
10        aProc.ActiveConnection = con
20        aProc.CommandText = "lims.lims_env.connect_same_session"
30        aProc.CommandType = adCmdStoredProc

40        aSession.Type = adDouble
50        aSession.Direction = adParamInput
60        aSession.Value = aSessionID
70        aProc.parameters.Append aSession

80        aProc.Execute
90        Set aSession = Nothing
100       Set aProc = Nothing
End Sub

Private Sub IExtensionWindow2_Close()
End Sub


Private Sub InitializeGrid()
          Dim x As Integer
          Dim s As String
              
10        grid.Clear
              
20        grid.AllowBigSelection = False
30        grid.AllowUserResizing = flexResizeNone
40        grid.Enabled = True
          
50        grid.ScrollBars = flexScrollBarBoth
60        grid.SelectionMode = flexSelectionFree
70        grid.AllowUserResizing = flexResizeBoth

80        grid.Font.Size = 12
90        grid.Rows = 2
100       grid.Cols = 7
110       grid.FixedRows = 1
120       grid.FixedCols = 0

130       grid.row = 0
140       grid.RowHeight(x) = 400
      '    For X = 1 To grid.Rows - 1
      '        grid.row = X
      '        grid.RowHeight(X) = 600
      '    Next X
150           grid.col = 0
160           grid.ColWidth(1) = 1400
170       For x = 1 To 3
180           grid.col = x
190           grid.ColWidth(x) = 2200
200       Next x

210       For x = 4 To grid.Cols - 1
220           grid.col = x
230           grid.ColWidth(x) = 2400
240       Next x
          
          'set the text for the COLUMN HEADERS:
250       grid.row = 0
260       grid.col = 0
      '    grid.CellAlignment = vbLeftJustify
      '    grid.Text = "Entity Type"

270       grid.CellAlignment = vbLeftJustify
280       grid.Text = "ID"
          
290       grid.col = grid.col + 1
300       grid.CellAlignment = vbLeftJustify
310       grid.Text = "Patho-Lab Name"

320       grid.col = grid.col + 1
330       grid.CellAlignment = vbLeftJustify
340       grid.Text = "Entity Name"

350       grid.col = grid.col + 1
360       grid.CellAlignment = vbLeftJustify
370       grid.Text = "Action"
          
380       grid.col = grid.col + 1
390       grid.CellAlignment = vbLeftJustify
400       grid.Text = "Details"
          
410       grid.col = grid.col + 1
420       grid.CellAlignment = vbLeftJustify
430       grid.Text = "Patholog" '"Created By"
          
440       grid.col = grid.col + 1
450       grid.CellAlignment = vbLeftJustify
460       grid.Text = "Created On"

End Sub

'used for a concatenated fieled: value1;value2
'gets the 1st value of that string
Private Function CleanSemicolon(str As String) As String
          Dim I As Integer
          
10        CleanSemicolon = str
          
20        I = InStr(1, str, ";")
30        If I = 0 Then Exit Function
          
40        CleanSemicolon = Left(str, I - 1)
End Function


Private Function nte(e As Variant) As Variant
10        nte = IIf(IsNull(e), "", e)
End Function


Private Sub lstEntityTypes_Click()
10        cmdShowRequests.Enabled = True
20        cmdShowRequests_Click
End Sub

'if a new request is made, of a type we do not currently
'have in the list of types (slide / block / sample etc.)
'it is added to the list:
Private Sub Timer1_Timer()
10        Timer1.Enabled = False
20        Call RefreshEntityTypes
30        Timer1.Enabled = True
End Sub

'mark the line(s) of this entity
'get the EXTRA_REQUSER_DATA_ID at all the entries this entity is found
'and add it to the dictionaty of barcoded entities
Private Sub txtEntityBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
10    On Error GoTo ERR_txtEntityBarcode_KeyDown
          Dim strEntityName As String

20        If KeyCode <> vbKeyReturn Then Exit Sub

30        strEntityName = UCase(txtEntityBarcode.Text)

40        Call BarcodeEntity(strEntityName)


50        txtEntityBarcode.Text = ""
          

60        Exit Sub
ERR_txtEntityBarcode_KeyDown:
70    MsgBox "ERR_txtEntityBarcode_KeyDown" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub BarcodeEntity(strEntityName As String)
10    On Error GoTo ERR_BarcodeEntity
          Dim I As Integer
          Dim k As Integer
          Dim s As String
          Dim MainStr As String
          Dim dicLocations As Dictionary



20        If Not dicEntities.Exists(strEntityName) Then
30            MsgBox "The barcode value doesn't exist in the list"
40            txtEntityBarcode.Text = ""
50            txtEntityBarcode.SetFocus
60            Exit Sub
70        End If

80        Set dicLocations = dicEntities(strEntityName)
      '    MainStr = dicEntities(strEntityName)
          
90        lblMicrotom.Visible = False
          
100       For I = 0 To dicLocations.Count - 1
110           s = dicLocations.Keys(I)
120           grid.row = CLng(s)
130           grid.col = 0

140           Call ShowMicrotom(entityType, grid.TextMatrix(grid.row, 2), _
                                                     grid.TextMatrix(grid.row, 0), _
                                                     strEntityName)

150           If Not dicBarcodeEntities.Exists(grid.Text) Then
160               Call dicBarcodeEntities.Add(grid.Text, strEntityName)
170           End If

180           For k = 0 To grid.Cols - 1
190               grid.col = k
200               grid.CellBackColor = MARK_SELECTED
210           Next k
220       Next I
          
   
          
230       cmdReport.Enabled = True

240       Exit Sub
ERR_BarcodeEntity:
250   MsgBox "ERR_BarcodeEntity" & vbCrLf & "In Line #" & Erl & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


'show the microtom for adding slide to the block if
'for the sub entity the DESCRIPTION is T
'(the doctor asked for it to be cut by the same microtom as the last one)
Private Sub ShowMicrotom(strEntityType As String, strAction As String, _
                         strExtraRequestDataId As String, _
                         strEntityName As String)
10    On Error GoTo ERR_ShowMicrotom
          Dim rs As Recordset
          Dim sql As String

20        If strEntityType <> "Block" Then Exit Sub
          
30        If strAction <> "Add Slide" Then Exit Sub
          
40        sql = " select rd.DESCRIPTION"
50        sql = sql & " from lims_sys.u_extra_request_data rd"
60        sql = sql & " where rd.U_EXTRA_REQUEST_DATA_ID='" & strExtraRequestDataId & "'"
              
70        Set rs = con.Execute(sql)
          
80        If rs.EOF Then Exit Sub

90        If nte(rs("DESCRIPTION")) = "T" Then
100           lblMicrotom.Caption = "Microtome: " & GetMicrotom(strEntityName)
110           lblMicrotom.Visible = True
120       End If
              
130       Exit Sub
ERR_ShowMicrotom:
140   MsgBox "ERR_ShowMicrotom" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Function getNextStr(ByRef s As String, c As String) As String
          Dim p
          Dim res
10        p = InStr(1, s, c)
20        If (p = 0) Then
30            res = s
40            s = ""
50            getNextStr = res
60        Else
70            res = Mid$(s, 1, p - 1)
80            s = Mid$(s, p + Len(c), Len(s))
90            getNextStr = res
100       End If
End Function

'debug-print the collection of all entities at the current selection
Private Sub DEBUG_SHOW_DICTIONARY()
10    On Error GoTo ERR_DEBUG_SHOW_DICTIONARY
          Dim I As Integer
          Dim j As Integer
          Dim s As String
          Dim MainStr As String
          

20        For I = 0 To dicEntities.Count - 1
30            s = "_" & dicEntities.Keys(I) & "_"
              
40            MainStr = dicEntities.Items(I)
              
50            While MainStr <> ""
60                s = s & vbCrLf & getNextStr(MainStr, ",")
70            Wend
              
80            MsgBox s
90        Next I

100       Exit Sub
ERR_DEBUG_SHOW_DICTIONARY:
110   MsgBox "ERR_DEBUG_SHOW_DICTIONARY" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub PrintToPrinter(iStartRow As Integer, iEndRow As Integer, iRows As Integer)
10    On Error GoTo ERR_PrintToPrinter

          Dim BottomOffset As Long
          Dim RightOffset As Long
          Dim strName As String
          Dim s As String
          


          Dim I As Integer, j As Integer   ' Declare variables
          
          
          
         ' NOT WORKING Printer.lpszDocName = "Extra_request_2"
20        Printer.ScaleMode = 3   ' Set ScaleMode to pixels.
30        Printer.Orientation = vbPRORLandscape
          
40        TopOffset = 400 'Int(Printer.ScaleHeight / 8)
50        BottomOffset = 50 'Int(Printer.ScaleHeight / 12)
60        LeftOffset = 100 'CInt(Printer.ScaleHeight / 12)
70        RightOffset = 50
80        ColSize = Int((Printer.ScaleWidth - LeftOffset - RightOffset) / (grid.Cols))
90        RowSize = Int((Printer.ScaleHeight - TopOffset - BottomOffset) / (iRows))
100       Printer.DrawWidth = 2   ' Set DrawWidth.
          
110       Printer.Font = "Arial"
120       Printer.FontSize = 10
130       Printer.FontBold = False
140       Printer.FontUnderline = False
150       Printer.CurrentY = Int(TopOffset / 2)
160       Printer.CurrentX = LeftOffset
          'to be modified:
          'Printer.Print "Plate Name: " & strPlateName '"Plate Name"
170       Printer.FontUnderline = False


          'print report name and dates above the table:
180       Printer.CurrentY = Int(TopOffset / 6)
190       Printer.CurrentX = Int(LeftOffset)
200       Printer.Print "Entity Type: " & lstEntityTypes.Text

210       Printer.FontSize = 8

          'print columns headers:
220       grid.row = 0
230       For I = 0 To grid.Cols - 1
240           grid.col = I
250           strName = Left(grid.Text, MAX_DIGITS_PER_CELL)
260         strName = Replace(strName, ">", "")
270         strName = Replace(strName, "<", "")
280           Printer.CurrentY = Int(TopOffset / 1.5)
              'Printer.CurrentY = 3 * Int(TopOffset / 4)
              
290           Printer.CurrentX = ColPixel(I)
              'Printer.CurrentX = ColPixel(i) + Int(ColSize / 2) - 50
              
300           Printer.Print strName
      '        While strName <> ""
      '            s = getNextStr(strName, vbCrLf)
      '            Printer.Print Left(s, 9)
      '            Printer.CurrentX = ColPixel(i)
      '            Printer.CurrentY = Printer.CurrentY + 10
      '
      '            If Len(s) > 9 Then
      '                Printer.Print Mid(s, 10, 9)
      '                Printer.CurrentX = ColPixel(i)
      '                Printer.CurrentY = Printer.CurrentY + 10
      '            End If
      '        Wend
310       Next I
          
          'print row headers:
      '    grid.col = 0
      '    For i = 0 To grid.Rows - 1
      '        grid.row = i
      '        strName = grid.Text
      '        Printer.CurrentX = Int(LeftOffset / 4)
      '        Printer.CurrentY = RowPixel(i) - 20
      '
      '        Printer.Print strName
      '
      ''        strName = Mid(strName, InStr(1, strName, " ", vbTextCompare) + 1)
      ''
      ''        While strName <> ""
      ''            s = getNextStr(strName, " ")
      ''            Printer.Print Left(s, 15)
      ''            Printer.CurrentX = Int(LeftOffset / 4)
      ''            Printer.CurrentY = Printer.CurrentY + 10
      ''        Wend
      '    Next i

          'try a diff font - fixed sized one
          'to do -
          '1. count how many digits can enter in a cell
          '2. change font size if needed:
          
320       strFontName = "Miriam Fixed"
330       iFontSize = 8
340       isFontBold = False
          
350       Printer.Font = strFontName
360       Printer.FontSize = iFontSize
370       Printer.FontBold = isFontBold

380       I = TopOffset
390       While I <= Printer.ScaleHeight - BottomOffset
400           Printer.Line (LeftOffset, I)-(Printer.ScaleWidth - RightOffset, I)
410           I = I + RowSize
420       Wend
430       I = LeftOffset
          

440       While I <= Int(Printer.ScaleWidth) + 1 - RightOffset
450           Printer.Line (I, TopOffset)-(I, Printer.ScaleHeight - BottomOffset)
460           I = I + ColSize
470       Wend

480       For I = 0 To grid.Cols - 1
490           grid.col = I
          
500           For j = iStartRow To iEndRow
                  Dim iNameSize As Integer
                  Dim iNameSizeInPixels As Integer
                  Dim iCenteredColPixel As Integer
                  Dim iCellLeftShift As Integer
                  
                  'change to udi report
                  '15.05.2006
                  'get the text from that cell in the grid
                  'strName=...
                  
510               grid.row = j
                  
520               strName = Left(grid.Text, MAX_DIGITS_PER_CELL)
   ' MsgBox strName
      '            strName = Replace(strName, vbCrLf, " ")
      '            strName = aAliquotArray(j, i).strName
530               iNameSize = Len(strName)
                  
      '            iNameSizeInPixels = iNameSize * ColSize / iMaxDigitsPerCell
                     
                  'in case the name is too long
                  'there in no shift at all:
      '            If iNameSizeInPixels > ColSize Then
      '                iNameSizeInPixels = ColSize
      '            End If
                  
                  'number of blank pixels from the left of the cell
                  'for this name:
540               iCellLeftShift = 0 ' (ColSize - iNameSizeInPixels) / 2
                  
                  'start printing this cell's name
                  'at that pixel:
            '      iCenteredColPixel = iCellLeftShift + ColPixel(i) '+ iNameSizeInPixels / 4
                  
            '      Printer.CurrentX = iCenteredColPixel
                  
550               Printer.CurrentX = ColPixel(I)
560               Printer.CurrentY = RowPixel(j - iStartRow) - 20
                  
570               Printer.Print strName
      '            While strName <> ""
      '                s = getNextStr(strName, vbCrLf)
      '                Printer.Print Left(s, 8)
      '                Printer.CurrentX = ColPixel(i)
      '                Printer.CurrentY = Printer.CurrentY + 10
      '            Wend
                  
                  
                  'not all the name is printed if there is not
                  'enough space:
                  
                  
                  'Printer.Print Left(strName, iMaxDigitsPerCell)
                  
                  
                  'Printer.Print Left(aAliquotArray(j, i), 11)
580           Next j
590       Next I
          'Printer.KillDoc
600       Printer.EndDoc
          
610       Exit Sub
ERR_PrintToPrinter:
620   MsgBox "ERR_PrintToPrinter" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub



Private Function ColPixel(col As Integer) As Long
10        ColPixel = LeftOffset + col * ColSize + 10
End Function


Private Function RowPixel(row As Integer) As Long
10        RowPixel = TopOffset + row * RowSize + Int(RowSize / 2)
          'RowPixel = TopOffset + row * RowSize + Int(RowSize / 2) - 50
End Function

'get list of entity types in status V
'(not yet reported as out of archive)
Private Sub InitEntityTypes()
10    On Error GoTo ERR_InitEntityTypes
          Dim rs As Recordset

20          If lstEntityTypes.ListCount = 0 Then
30                lstEntityTypes.AddItem (IM)
40               lstEntityTypes.AddItem (HI)
50              End If
          
60        If lstEntityTypes.ListCount > 0 Then
70            lstEntityTypes.Selected(0) = True
80            cmdShowRequests.Enabled = True
90            Call cmdShowRequests_Click
100       Else
110           cmdShowRequests.Enabled = False
120       End If

130       Exit Sub
ERR_InitEntityTypes:
140   MsgBox "ERR_InitEntityTypes" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

'add to the list of entity types:
Private Sub RefreshEntityTypes()
10    On Error GoTo ERR_RefreshEntityTypes
         
          
20      If lstEntityTypes.ListCount = 0 Then
30                lstEntityTypes.AddItem (IM)
40               lstEntityTypes.AddItem (HI)
50              End If
            

60        Exit Sub
ERR_RefreshEntityTypes:
70    MsgBox "ERR_RefreshEntityTypes" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

'get the letter representing the last workstation to use
'SlideGenerationReport for this block:
Private Function GetMicrotom(strBlockName As String)
10    On Error GoTo ERR_GetMicrotom
          Dim sql As String
          Dim rs As Recordset
          Dim strWorkstationId As String
          Dim strWorkstationName As String

          'get the workstation id:
20        sql = " select au.U_LAST_MICROTOME"
30        sql = sql & " from lims_sys.aliquot a,lims_sys.aliquot_user au"
40        sql = sql & " where a.ALIQUOT_ID=au.ALIQUOT_ID"
50        sql = sql & " and a.NAME='" & strBlockName & "'"

60        Set rs = con.Execute(sql)
70        strWorkstationId = nte(rs("U_LAST_MICROTOME"))
80        If strWorkstationId = "" Then Exit Function
          
          'get the workstation name:
90        sql = " select name "
100       sql = sql & " from lims_sys.workstation "
110       sql = sql & " where workstation_id = " & strWorkstationId
          
120       Set rs = con.Execute(sql)
130       strWorkstationName = rs("name")

          'get the microtom for the workstation name:
140       Set rs = con.Execute("select phrase_description from lims_sys.phrase_entry " & _
              "where phrase_name = '" & strWorkstationName & "' and " & _
              "phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
              "name = 'StationToMicrotom')")
              
150       If Not rs.EOF Then
160           GetMicrotom = nte(rs("phrase_description"))
170       End If
          
180       Exit Function
ERR_GetMicrotom:
190   MsgBox "ERR_GetMicrotom" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function

'a click on a cell shows the remark for the request
'containing this entity
Private Sub grid_DblClick()
10    On Error GoTo ERR_grid_Click
          Dim strRequestDataId As String
          
20        strRequestDataId = grid.TextMatrix(grid.row, 0)

30        Call frmRemarks.Initialize(con, strRequestDataId)
40        Call frmRemarks.Show(vbModal)
          
50        Exit Sub
ERR_grid_Click:
60    MsgBox "MsgBox" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description

End Sub

'ashi Sorting
Private Sub grid_MouseUp(Button As Integer, Shift As _
    Integer, x As Single, y As Single)
          ' If this is not row 0, do nothing.
10        If grid.MouseRow <> 0 Then Exit Sub

          ' Sort by the clicked column.
20        SortByColumn grid.MouseCol
End Sub

' Sort by the indicated column.
Private Sub SortByColumn(ByVal sort_column As Integer)
          ' Hide the FlexGrid.
10        grid.Visible = False
20        grid.Refresh

          ' Sort using the clicked column.
30        grid.col = sort_column
40        grid.ColSel = sort_column
50        grid.row = 0
60        grid.RowSel = 0

          ' If this is a new sort column, sort ascending.
          ' Otherwise switch which sort order we use.
70        If m_SortColumn <> sort_column Then
80            m_SortOrder = flexSortGenericAscending
90        ElseIf m_SortOrder = flexSortGenericAscending Then
100           m_SortOrder = flexSortGenericDescending
110       Else
120           m_SortOrder = flexSortGenericAscending
130       End If
140       grid.Sort = m_SortOrder

          ' Restore the previous sort column's name.
150       If m_SortColumn >= 0 Then
160           grid.TextMatrix(0, m_SortColumn) = _
                  Mid$(grid.TextMatrix(0, m_SortColumn), 3)
170       End If

          ' Display the new sort column's name.
180       m_SortColumn = sort_column
190       If m_SortOrder = flexSortGenericAscending Then
200           grid.TextMatrix(0, m_SortColumn) = "> " & _
                  grid.TextMatrix(0, m_SortColumn)
210       Else
220           grid.TextMatrix(0, m_SortColumn) = "< " & _
                  grid.TextMatrix(0, m_SortColumn)
230       End If

          ' Display the FlexGrid.
240       grid.Visible = True
End Sub

Private Function ExistRemark(strExtraRequestDataId As String) As Boolean
10    On Error GoTo ERR_ExistRemark
          Dim rs As Recordset
          Dim sql As String
          
20        sql = " select r.DESCRIPTION "
30        sql = sql & " from lims_sys.u_extra_request_data rd, "
40        sql = sql & "      lims_sys.u_extra_request_data_user rdu, "
50        sql = sql & "      lims_sys.u_extra_request r"
60        sql = sql & "  where rd.U_EXTRA_REQUEST_DATA_ID=rdu.U_EXTRA_REQUEST_DATA_ID"
70        sql = sql & "  and   r.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
80        sql = sql & "  and   rd.U_EXTRA_REQUEST_DATA_ID=" & strExtraRequestDataId

90        Set rs = con.Execute(sql)
          
100       If nte(rs("DESCRIPTION")) = "" Then
110           ExistRemark = False
120       Else
130           ExistRemark = True
140       End If

150       Exit Function
ERR_ExistRemark:
160   MsgBox "ERR_ExistRemark" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function


'get the shortcut for the color group of this color
'an empty string is returned if this is not on of the known colors
Private Function GetColorGroup(strColor As String) As String
10    On Error GoTo ERR_GetColorGroup
          Dim rs As Recordset
          Dim sql As String
          Dim s As String
          
20        sql = " select ph.NAME "
30        sql = sql & " from lims_sys.phrase_header ph,"
40        sql = sql & "      lims_sys.phrase_entry pe"
50        sql = sql & " where pe.PHRASE_ID=ph.PHRASE_ID"
60        sql = sql & " and pe.PHRASE_NAME='" & strColor & "' "
70        sql = sql & " and ph.NAME in ('Pathology Molecular Stains',"
80        sql = sql & "                 'Pathology Special Stains',"
90        sql = sql & "                 'Pathology Other Stain Options',"
100       sql = sql & "                 'Pathology Imonohistochemistry stains')"

110       Set rs = con.Execute(sql)
          
120       If rs.EOF = True Then Exit Function
          
130       s = nte(rs("NAME"))
          
140       Select Case s
              Case "Pathology Molecular Stains"
150               s = "Mol"
160           Case "Pathology Special Stains"
170               s = "S"
180           Case "Pathology Imonohistochemistry stains"
190               s = "IHC"
200           Case "Pathology Other Stain Options"
210               s = "O"
220       End Select

230       GetColorGroup = s

240       Exit Function
ERR_GetColorGroup:
250   MsgBox "ERR_GetColorGroup" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function

'indicate if this item (sample / block / slide)
'is in the tissue archive or not:
Private Sub UpdateArchive(strEntityType As String, strEntityName As String, strStored As String)
10    On Error GoTo ERR_UpdateArchive
          Dim sql As String
              
20        Select Case strEntityType
              Case "Sample"
30                sql = " update lims_sys.sample_user su"
40                sql = sql & " set su.U_ARCHIVE='" & strStored & "'"
50                sql = sql & " where su.SAMPLE_ID="
60                sql = sql & " ("
70                sql = sql & "   select s.SAMPLE_ID"
80                sql = sql & "   from lims_sys.sample s"
90                sql = sql & "   where s.NAME='" & strEntityName & "'"
100               sql = sql & " )"
              
110           Case "Block"
120               sql = " update lims_sys.aliquot_user au"
130               sql = sql & " set au.U_ARCHIVE='" & strStored & "'"
140               sql = sql & " where au.ALIQUOT_ID="
150               sql = sql & " ("
160               sql = sql & "   select a.ALIQUOT_ID"
170               sql = sql & "   from lims_sys.aliquot a"
180               sql = sql & "   where a.NAME='" & strEntityName & "'"
190               sql = sql & " )"
200               sql = sql & " and exists"
210               sql = sql & " ("
220               sql = sql & "   select 1 "
230               sql = sql & "   from lims_sys.aliquot_formulation af"
240               sql = sql & "   where af.PARENT_ALIQUOT_ID=au.ALIQUOT_ID"
250               sql = sql & " )"
              
260           Case Else
270               sql = " update lims_sys.aliquot_user au"
280               sql = sql & " set au.U_ARCHIVE='" & strStored & "'"
290               sql = sql & " where au.ALIQUOT_ID="
300               sql = sql & " ("
310               sql = sql & "   select a.ALIQUOT_ID"
320               sql = sql & "   from lims_sys.aliquot a"
330               sql = sql & "   where a.NAME='" & strEntityName & "'"
340               sql = sql & " )"
350               sql = sql & " and exists"
360               sql = sql & " ("
370               sql = sql & "   select 1 "
380               sql = sql & "   from lims_sys.aliquot_formulation af"
390               sql = sql & "   where af.CHILD_ALIQUOT_ID=au.ALIQUOT_ID"
400               sql = sql & " )"
                  
410       End Select
          
          

          
420       Call con.Execute(sql)

430       Exit Sub
ERR_UpdateArchive:
440   MsgBox "UpdateArchive" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

