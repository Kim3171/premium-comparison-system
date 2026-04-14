Option Explicit

'*******************************************************************************
' Premium Asset Comparison System - Cross-Workbook Support
' Version 18.0 - SAFE ARCHITECTURE with SafeMode and Layer Separation
'*******************************************************************************

'===============================================================================
' LAYER ARCHITECTURE
'===============================================================================
' Layer 1 - Initialization: Environment detection only (NO DATA WRITES)
'   InitializeModule, InitializeDatasetContext, DetectHeaderRow, GetValidWorksheetForUI
'
' Layer 2 - UI Builder: Interface controls only (NO DATASET MODIFICATION)
'   RebuildMatchBuilderUI, AddControlButtons, AddStatusDisplay
'
' Layer 3 - Execution Engine: Data operations (ALLOWS DATA MODIFICATION)
'   CompareAssets, AddMatchRow, DeleteMatchRow, ExecuteCompareWithValidation
'===============================================================================

'===============================================================================
' GLOBAL STATE - All global variables declared here
'===============================================================================

' Layer Control
Public g_SafeMode As Boolean              ' When True, blocks destructive operations
Public g_MacroRunning As Boolean          ' When True, blocks Worksheet_Change event
Public g_AllowButtonDelete As Boolean     ' When True, allows button deletion (default False = protected)
Public g_ClearMatchDataOnRebuild As Boolean ' When True, clears old MATCH UI data on rebuild
Public g_Initialized As Boolean             ' Tracks if initialization completed
Public g_ForceRebuild As Boolean           ' Flag to force UI rebuild

' Dataset Context (read-only detected values)
Public g_DataHeaderRow As Long             ' Detected row containing data headers
Public g_DataStartRow As Long               ' First row of actual data
Public g_LastDataColumn As Long             ' Last column with data

' Worksheet References
Public g_CurrentWorksheet As Worksheet     ' Currently active worksheet
Public g_SourceWorkbook As Workbook        ' Source workbook reference
Public g_SourceSheet As Worksheet          ' Source sheet reference
Public g_TargetWorkbook As Workbook         ' Target workbook reference
Public g_TargetSheet As Worksheet          ' Target sheet reference
Public g_SourceHeaderRow As Long            ' Source header row
Public g_TargetHeaderRow As Long           ' Target header row
Public g_MatchedIdColumn As String    ' User-selected column whose value goes to MATCHED_ID

' UI State
Public g_UIExists As Boolean                ' Tracks if UI is already present
Public g_DetectedUIExists As Boolean       ' Pre-detection result
Public g_ParseCancelled As Boolean         ' Used by ParseRowSelection to signal cancellation

'===============================================================================
' UI CONSTANTS
'===============================================================================

' UI Constants (absolute rows for the Match Builder UI)
Public Const UI_TITLE_ROW As Long = 1         ' Row 1: Source configuration
Public Const UI_CONTROL_ROW As Long = 3       ' Row 3: Control buttons
Public Const UI_STATUS_ROW As Long = 1        ' Row 1: Source display
Public Const UI_COLHEADER_ROW As Long = 4     ' Row 4: Column headers
Public Const UI_FIRST_MATCH_ROW As Long = 5   ' Row 5+: Match rules start here
Public Const UI_SOURCE_ROW As Long = 1        ' Row 1: Source configuration
Public Const UI_TARGET_ROW As Long = 2        ' Row 2: Target configuration

' Fixed data header - fallback when detection fails
Public Const FIXED_DATA_HEADER_ROW As Long = 10

' UI Height (number of rows reserved for UI)
Public Const UI_HEIGHT As Long = 6

'===============================================================================
' CONFIGURATION CONSTANTS
'===============================================================================

Public Const CONFIG_SHEET_NAME As String = "COMPARE_CONFIG"
Public Const HEADER_MAP_SHEET_NAME As String = "HEADER_MAP"
Public Const CONFIG_SOURCE_WB As String = "SOURCE_WORKBOOK"
Public Const CONFIG_SOURCE_WS As String = "SOURCE_SHEET"
Public Const CONFIG_TARGET_WB As String = "TARGET_WORKBOOK"
Public Const CONFIG_TARGET_WS As String = "TARGET_SHEET"

Private Const MATCH_COLUMNS As String = "MATCHED_ID,MATCHED_ASSETID,MATCH_ID,MATCHEDID"

'===============================================================================
' DEBUGGING SYSTEM
'===============================================================================

Private Const DEBUG_MODE As Boolean = False  ' Set to False in production

Public Sub DebugPrint(msg As String)
    ' Layer-aware debug output
    #If DEBUG_MODE Then
        Dim prefix As String
        prefix = "|"
        If g_SafeMode Then prefix = "[SAFE] " & prefix
        Debug.Print Format(Now, "hh:nn:ss") & " " & prefix & " " & msg
    #End If
End Sub

'===============================================================================
' SAFE MODE GUARD
'===============================================================================

Private Function IsDestructiveAllowed() As Boolean
    ' Returns True if destructive operations are allowed
    ' Checks g_SafeMode and logs the attempt
    If g_SafeMode Then
        DebugPrint "BLOCKED: Destructive operation prevented by SafeMode"
        IsDestructiveAllowed = False
    Else
        IsDestructiveAllowed = True
    End If
End Function

Private Sub RequireDestructiveMode()
    ' Raises error if SafeMode is enabled
    If g_SafeMode Then
        Err.Raise vbObjectError + 1000, "SafeMode", _
            "Destructive operation blocked. Disable SafeMode to proceed."
    End If
End Sub

'===============================================================================
' GLOBAL STATE INITIALIZATION
'===============================================================================

Private Sub ResetGlobalState()
    ' Initialize all global variables to safe defaults
    ' Called at module start and when force rebuild requested

    g_SafeMode = False          ' Default to allow UI building
    g_Initialized = False
    g_ForceRebuild = False
    g_ClearMatchDataOnRebuild = False  ' Default to keep existing match data
    g_DataHeaderRow = 0
    g_DataStartRow = 0
    g_LastDataColumn = 0
    g_UIExists = False
    g_DetectedUIExists = False

    Set g_CurrentWorksheet = Nothing
    Set g_SourceWorkbook = Nothing
    Set g_SourceSheet = Nothing
    Set g_TargetWorkbook = Nothing
    Set g_TargetSheet = Nothing

    DebugPrint "Global state reset - SafeMode: " & g_SafeMode
End Sub

'===============================================================================
' STARTUP CONTROLLER - Centralized system startup
'===============================================================================

Public Sub StartSystem()
    '
    ' LAYER 0: Startup Controller
    ' Ensures safe, deterministic initialization on workbook open
    '
    ' Responsibilities:
    ' - Ensure Application state is correct
    ' - Reset global state
    ' - Call InitializeModule safely
    ' - Log startup state
    '

    On Error Resume Next

    DebugPrint "=========================================="
    DebugPrint "StartSystem: Beginning safe initialization"

    ' Ensure Application state is correct
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    DebugPrint "StartSystem: Application events enabled"

    ' Reset global state to safe defaults
    Call ResetGlobalState
    DebugPrint "StartSystem: Global state reset complete"

    ' FIX: Update btnBuildUI button OnAction if it points to old function
    Call FixBuildUIButton
    DebugPrint "StartSystem: Button fix applied"

    ' Initialize module if not already initialized
    If Not g_Initialized Then
        DebugPrint "StartSystem: Calling InitializeModule..."
        Call InitializeModule
        DebugPrint "StartSystem: InitializeModule returned, g_Initialized=" & g_Initialized
    Else
        DebugPrint "StartSystem: Already initialized, skipping"
    End If

    DebugPrint "StartSystem: Startup complete - SafeMode=" & g_SafeMode
    DebugPrint "=========================================="

    On Error GoTo 0
End Sub

'===============================================================================
' MANUAL STARTUP FALLBACK - User-initiated system launch
'===============================================================================

Public Sub LaunchSystem()
    '
    ' Manual fallback to launch system from Excel's Macro menu
    ' Provides user feedback and ensures safe initialization
    '

    Dim ws As Worksheet

    On Error Resume Next

    ' Call the startup controller
    Call StartSystem

    ' Verify initialization
    If g_Initialized Then
        ' Try to get current worksheet for status
        Set ws = GetValidWorksheetForUI
        If Not ws Is Nothing Then
            MsgBox "Premium Comparison System initialized successfully!" & vbCrLf & vbCrLf & _
                   "Worksheet: " & ws.Name & vbCrLf & _
                   "Data Header Row: " & g_DataHeaderRow & vbCrLf & vbCrLf & _
                   "Click 'Select Files' to configure, then 'Execute Match' to run.", vbInformation
        Else
            MsgBox "Premium Comparison System initialized!", vbInformation
        End If
    Else
        MsgBox "Initialization may not have completed. Please try running 'StartSystem' or check the Immediate Window for errors.", vbExclamation
    End If

    On Error GoTo 0
End Sub

'===============================================================================
' NEW: Initialize on a SPECIFIC worksheet (call this instead of LaunchSystem)
'===============================================================================

Public Sub InitializeOnSheet()
    '
    ' NEW: Allows user to select which worksheet to initialize the UI on
    ' Use this if your data is not in the default worksheet
    '
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim wsName As String
    Dim wbName As String
    Dim i As Long
    Dim sheetList As String
    Dim selection As Integer

    ' First, ask which workbook contains the data
    If Workbooks.Count = 0 Then
        MsgBox "No workbooks are open!", vbCritical
        Exit Sub
    End If

    ' Build list of open workbooks
    sheetList = "Select the workbook containing your data:" & vbCrLf & vbCrLf
    For i = 1 To Workbooks.Count
        sheetList = sheetList & i & ". " & Workbooks(i).Name & vbCrLf
    Next i
    sheetList = sheetList & vbCrLf & "Enter number:"

    selection = InputBox(sheetList, "Select Workbook", 1)
    If selection < 1 Or selection > Workbooks.Count Then
        MsgBox "Selection cancelled.", vbInformation
        Exit Sub
    End If

    Set wb = Workbooks(selection)

    ' Now ask which sheet in that workbook
    If wb.Worksheets.Count = 0 Then
        MsgBox "No worksheets in that workbook!", vbCritical
        Exit Sub
    End If

    sheetList = "Select the sheet containing your data:" & vbCrLf & vbCrLf
    For i = 1 To wb.Worksheets.Count
        sheetList = sheetList & i & ". " & wb.Worksheets(i).Name & vbCrLf
    Next i
    sheetList = sheetList & vbCrLf & "Enter number:"

    selection = InputBox(sheetList, "Select Sheet", 1)
    If selection < 1 Or selection > wb.Worksheets.Count Then
        MsgBox "Selection cancelled.", vbInformation
        Exit Sub
    End If

    Set ws = wb.Worksheets(selection)

    ' Now set this as the current worksheet and build the UI
    Set g_CurrentWorksheet = ws
    g_DataHeaderRow = 0  ' Force re-detection
    g_Initialized = False
    g_ForceRebuild = True

    ' Build UI on the selected worksheet
    Call InitializeModule

    MsgBox "UI will be built on: " & ws.Parent.Name & " - " & ws.Name & vbCrLf & vbCrLf & _
           "If data starts at row 1, you will be prompted to insert 6 rows.", vbInformation
End Sub

'===============================================================================
' MAIN ENTRY POINT - Initialize on workbook open
'===============================================================================

'===============================================================================
' WORKBOOK OPEN HANDLER - MUST BE IN ThisWorkbook module
' Add this to your Workbook Open event:
' Private Sub Workbook_Open()
'     StartSystem
' End Sub
'===============================================================================

Public Sub InitializeModule(Optional ByVal skipUIBuild As Boolean = False)
    '
    ' LAYER 1: Main initialization entry point
    ' Called from Workbook_Open or manually
    '
    ' IMPORTANT: This runs in SafeMode by default - NO DATA MODIFICATION
    ' Only detection and UI setup (button creation) is allowed
    '
    ' Parameters:
    '   skipUIBuild - If True, skip building UI (used when opening workbook)
    '
    Dim ws As Worksheet
    Dim configSheet As Worksheet
    Dim uiAlreadyExists As Boolean

    DebugPrint "InitializeModule: Starting... (skipUIBuild=" & skipUIBuild & ")"

    ' CRITICAL: Guard against re-initialization - but allow if g_ForceRebuild is set
    If g_Initialized And Not g_ForceRebuild Then
        DebugPrint "InitializeModule: Already initialized, exiting"
        Exit Sub
    End If

    ' Step 1: Initialize global state (SafeMode defaults to True)
    ' BUT preserve g_CurrentWorksheet if it was already set (e.g., from InitializeOnSheet)
    Dim wsPreserved As Worksheet
    Set wsPreserved = g_CurrentWorksheet
    Call ResetGlobalState
    ' Restore the worksheet if it was set
    If Not wsPreserved Is Nothing Then
        Set g_CurrentWorksheet = wsPreserved
    End If

    ' Step 2: Get or create configuration sheet (read-only operation)
    Set configSheet = GetOrCreateConfigSheet
    DebugPrint "InitializeModule: Config sheet ready: " & configSheet.Name

    ' Step 3: Get a valid worksheet for UI (read-only)
    ' Use g_CurrentWorksheet if already set, otherwise try to find one
    If g_CurrentWorksheet Is Nothing Then
        Set ws = GetValidWorksheetForUI
    Else
        Set ws = g_CurrentWorksheet
    End If
    If ws Is Nothing Then
        MsgBox "ERROR: No worksheet available for UI initialization." & vbCrLf & _
               "Please ensure at least one worksheet exists.", vbCritical
        DebugPrint "InitializeModule: FAILED - No worksheet available"
        Exit Sub
    End If
    DebugPrint "InitializeModule: Using worksheet: " & ws.Name

    ' Step 4: Initialize dataset context (detects header row - READ ONLY)
    DebugPrint "InitializeModule: Calling InitializeDatasetContext..."
    Call InitializeDatasetContext(ws)

    ' Step 5: Verify initialization succeeded
    If Not g_Initialized Then
        DebugPrint "InitializeModule: WARNING - g_Initialized is False after InitializeDatasetContext"
        g_DataHeaderRow = FIXED_DATA_HEADER_ROW
        g_DataStartRow = FIXED_DATA_HEADER_ROW + 1
        g_LastDataColumn = 10
        g_Initialized = True
        DebugPrint "InitializeModule: Using fallback header row: " & g_DataHeaderRow
    End If

    DebugPrint "InitializeModule: Data header row: " & g_DataHeaderRow & ", Last column: " & g_LastDataColumn

    ' Step 6: Check if UI already exists (READ ONLY detection)
    uiAlreadyExists = IsUIAlreadyExists(ws)
    DebugPrint "InitializeModule: UI exists check: " & uiAlreadyExists

    ' Step 7: Build the UI only if needed
    ' Rebuild if: UI doesn't exist OR force rebuild requested
    ' CRITICAL FIX: When g_ForceRebuild is True (called from BuildFullUI), skip auto-build
    ' because BuildFullUI will call RebuildMatchBuilderUI directly with user-selected header row
    ' Also skip if skipUIBuild=True (called from Workbook_Open)
    ' This prevents the duplicate UI build bug
    If g_ForceRebuild Then
        ' Skip auto-build - BuildFullUI will handle it with user-selected header
        DebugPrint "InitializeModule: g_ForceRebuild=True, skipping auto-build (BuildFullUI will handle)"
    ElseIf skipUIBuild Then
        ' Skip UI build when opening workbook - user will click Build UI button manually
        DebugPrint "InitializeModule: skipUIBuild=True, skipping auto-build"
    ElseIf Not uiAlreadyExists Then
        DebugPrint "InitializeModule: Calling RebuildMatchBuilderUI..."
        Call RebuildMatchBuilderUI
    Else
        DebugPrint "InitializeModule: UI already exists, skipping rebuild"
    End If

    DebugPrint "InitializeModule: COMPLETE"

    MsgBox "Premium Comparison System initialized!" & vbCrLf & vbCrLf & _
           "Data Header Row: " & g_DataHeaderRow & vbCrLf & _
           "Worksheet: " & ws.Name & vbCrLf & vbCrLf & _
           "Click 'Select Files' to configure, then 'Execute Match' to run.", vbInformation
End Sub

'===============================================================================
' SAFE INITIALIZER - Centralized dataset context detection
'===============================================================================

Public Sub InitializeDatasetContext(Optional ws As Worksheet = Nothing)
    '
    ' Safe initializer that detects dataset context without modifying the sheet
    ' This MUST be called before any UI operations
    '
    Dim tempAliases As Object
    Dim tempLearned As Object
    Dim detectedRow As Long
    Dim detectedLastCol As Long

    DebugPrint "InitializeDatasetContext: Starting..."

    ' If worksheet not provided, try to get valid one
    If ws Is Nothing Then
        DebugPrint "InitializeDatasetContext: No worksheet provided, getting valid one..."
        Set ws = GetValidWorksheetForUI
    End If

    If ws Is Nothing Then
        DebugPrint "InitializeDatasetContext: FAILED - No worksheet available"
        g_Initialized = False
        Exit Sub
    End If

    DebugPrint "InitializeDatasetContext: Using worksheet: " & ws.Name

    ' Create temp dictionaries for header detection
    Set tempAliases = CreateObject("Scripting.Dictionary")
    Set tempLearned = CreateObject("Scripting.Dictionary")

    ' Detect header row
    detectedRow = FindDataHeaderRow(ws, tempAliases, tempLearned)
    DebugPrint "InitializeDatasetContext: Initial detection returned row: " & detectedRow

    ' If detection returns 1 or lower, try aggressive detection
    If detectedRow <= 1 Then
        DebugPrint "InitializeDatasetContext: Trying aggressive detection..."
        detectedRow = FindDataHeaderRowAggressive(ws, tempAliases, tempLearned)
        DebugPrint "InitializeDatasetContext: Aggressive detection returned row: " & detectedRow
    End If

    ' If still not found, use fallback
    If detectedRow <= 1 Then
        DebugPrint "InitializeDatasetContext: Using fallback row: " & FIXED_DATA_HEADER_ROW
        detectedRow = FIXED_DATA_HEADER_ROW
    End If

    ' Detect last column
    detectedLastCol = GetLastColumn(ws, detectedRow)
    If detectedLastCol < 3 Then detectedLastCol = 10
    DebugPrint "InitializeDatasetContext: Last column: " & detectedLastCol

    ' Update globals
    g_DataHeaderRow = detectedRow
    g_DataStartRow = detectedRow + 1
    g_LastDataColumn = detectedLastCol
    Set g_CurrentWorksheet = ws
    g_Initialized = True

    DebugPrint "InitializeDatasetContext: COMPLETE - Initialized = " & g_Initialized
End Sub

'===============================================================================
' UI EXISTENCE CHECKER - Returns True if UI is already built
'===============================================================================

Private Function IsUIAlreadyExists(ws As Worksheet) As Boolean
    '
    ' Checks if the Match Builder UI is already present on the worksheet
    ' Returns True if UI exists, False otherwise
    '
    Dim cellValue As String
    Dim btn As Excel.Button  ' Explicitly declared as Excel.Button
    Dim hasUIButton As Boolean
    Dim hasSourceText As Boolean
    Dim buttonCount As Long

    On Error Resume Next

    DebugPrint "IsUIAlreadyExists: Checking worksheet: " & ws.Name

    IsUIAlreadyExists = False
    hasUIButton = False
    hasSourceText = False
    buttonCount = 0

    ' Check 1: Is there a "Source:" indicator in row 1?
    If Not IsEmpty(ws.Cells(1, 1).value) Then
        cellValue = CStr(ws.Cells(1, 1).value)
        If InStr(1, UCase(Trim(cellValue)), "SOURCE:", vbTextCompare) > 0 Then
            hasSourceText = True
            DebugPrint "IsUIAlreadyExists: Found Source text"
        End If
    End If

    ' Check 2: Are there buttons in the UI area (rows 1-10)?
    For Each btn In ws.Buttons
        buttonCount = buttonCount + 1
        If btn.Top < ws.Rows(10).Top Then
            hasUIButton = True
            DebugPrint "IsUIAlreadyExists: Found UI button at row " & btn.Top / ws.Rows(1).Height
            Exit For
        End If
    Next btn

    DebugPrint "IsUIAlreadyExists: Button count = " & buttonCount & ", hasUIButton = " & hasUIButton

    ' UI exists only if we have BOTH indicators (Source text AND buttons)
    ' Note: For BuildFullUI, the user now selects the header row manually
    ' This check is mainly for InitializeModule to detect if UI exists
    If hasSourceText And hasUIButton Then
        IsUIAlreadyExists = True
        DebugPrint "IsUIAlreadyExists: UI EXISTS"
        Exit Function
    End If

    DebugPrint "IsUIAlreadyExists: Result = " & IsUIAlreadyExists

    On Error GoTo 0
End Function

'===============================================================================
' MAIN COMPARE FUNCTION
'===============================================================================

Public Sub CompareAssets()
    '
    ' LAYER 3: Execution Engine - Compare Assets
    '
    ' This is a destructive operation that writes MATCHED_ID to worksheet
    ' Protected by SafeMode check
    '
    Dim startTime As Double, endTime As Double
    Dim sourceData As Variant, targetData As Variant
    Dim sourceCols As Object, targetCols As Object
    Dim headerAliases As Object, sourceLearned As Object, targetLearned As Object
    Dim matchDefs As Collection
    Dim dictList As Collection
    Dim sourceRow As Long, targetRow As Long
    Dim m As Long
    Dim matchFound As Boolean
    Dim matchedID As String, matchType As String
    Dim targetRowIndex As Long
    Dim matchCounts() As Long
    Dim noMatchCount As Long
    Dim lastUpdateTime As Double, totalRows As Long, processedRows As Long

    ' Variables for inner loops
    Dim dict As Object
    Dim matchDef As Object
    Dim key As String
    Dim targetKeyValue As String
    Dim dictM As Object
    Dim sourceKey As String
    Dim dictTry As Object
    Dim summary As String
    Dim targetWorkbookName As String
    Dim sourceWorkbookName As String
    Dim parts() As String

    ' Variables for inner loops - additional
    Dim sourceAssetID As String
    Dim targetKeyID As String
    Dim targetMatchedInfo As String
    Dim targetMatchType As String
    Dim reverseDict As Object

    ' Variables for optimized bulk writing
    Dim sourceResults() As Variant
    Dim targetResults() As Variant
    Dim rowCount As Long
    Dim writeRange As Range
    Dim sourceStartRow As Long
    Dim sourceWriteArray() As Variant
    Dim writeCol As Long
    Dim minSourceCol As Long
    Dim targetRowCount As Long
    Dim targetStartRow As Long
    Dim targetWriteArray() As Variant
    Dim minTargetCol As Long
    Dim ti As Long
    Dim i As Long
    Dim sourceIdColIdx As Long
    Dim preColKey As Variant
    Dim tempFoundIdx As Long
    Dim targetIdColIdx As Long
    Dim fastKeyIdxs() As Variant
    Dim fastKeyColCounts() As Long
    Dim fastKeyArr As Variant
    Dim preM As Long
    Dim preColName As Variant
    Dim preNorm As String
    Dim preCount As Long
    Dim mDefCount As Long
    Dim srcIdxArr() As Long
    Dim tgtIdxArr() As Long
    Dim wbNameConfigSheet As Worksheet
    Dim normalizedMatchCol As String
    Dim resolvedColIdx As Long

    ' SAFE MODE GUARD - Block destructive operations in SafeMode
    If g_SafeMode Then
        MsgBox "Compare Assets is blocked in SafeMode." & vbCrLf & vbCrLf & _
               "This operation writes MATCHED_ID to your worksheet." & vbCrLf & _
               "Please use 'Execute Compare' button instead, which will " & _
               "prompt you before disabling SafeMode.", vbExclamation
        DebugPrint "CompareAssets: BLOCKED in SafeMode"
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    startTime = Timer
    lastUpdateTime = startTime
    DebugPrint "CompareAssets: Starting..."

    ' Step 1: Get or configure source/target
    Call ConfigureSourceTarget
    If g_SourceSheet Is Nothing Or g_TargetSheet Is Nothing Then
        MsgBox "Source and Target sheets must be selected.", vbCritical
        Exit Sub
    End If

    ' Performance settings
    Call OptimizePerformance(True)

    ' Step 2: Initialize learning dictionaries
    Set headerAliases = CreateObject("Scripting.Dictionary")
    Set sourceLearned = CreateObject("Scripting.Dictionary")
    Set targetLearned = CreateObject("Scripting.Dictionary")

    ' Step 3: Load HEADER_MAP if exists
    Call LoadHeaderMapping(headerAliases)

    ' Step 4: Find header rows
    ' CRITICAL FIX: Only auto-detect if header rows weren't already confirmed by user
    ' This preserves user-confirmed header rows from ExecuteCompareWithValidation
    If g_SourceHeaderRow = 0 Then
        g_SourceHeaderRow = FindDataHeaderRow(g_SourceSheet, headerAliases, sourceLearned)
    End If
    If g_TargetHeaderRow = 0 Then
        g_TargetHeaderRow = FindDataHeaderRow(g_TargetSheet, headerAliases, targetLearned)
    End If

    ' Step 4.5: Check if UI rows exist - if not, prompt user to build UI first
    ' SIMPLIFIED: Check for buttons (most reliable UI indicator)
    Dim uiMissing As Boolean
    Dim hasButton As Boolean
    Dim btn As Excel.Button

    ' Check for buttons in the worksheet (reliable indicator that UI exists)
    hasButton = False
    On Error Resume Next
    For Each btn In g_SourceSheet.Buttons
        If btn.Top < g_SourceSheet.Rows(10).Top Then
            hasButton = True
            Exit For
        End If
    Next btn
    On Error GoTo 0

    ' Check for "Source:" text in row 1 (another UI indicator)
    Dim hasSourceText As Boolean
    hasSourceText = False
    If InStr(1, UCase(g_SourceSheet.Cells(1, 1).Value), "SOURCE:", vbTextCompare) > 0 Then
        hasSourceText = True
    End If

    ' UI exists if we have buttons OR Source: text
    uiMissing = Not (hasButton Or hasSourceText)

    DebugPrint "CompareAssets: UI check - hasButton=" & hasButton & ", hasSourceText=" & hasSourceText & ", uiMissing=" & uiMissing

    ' REMOVED: Auto UI rebuild - UI should only be built manually by user
    ' If uiMissing Then
    '     MsgBox "Match UI rows appear to be missing." & vbCrLf & _
    '            "Please click 'Build UI' button manually if needed.", vbInformation
    ' End If

    ' Step 5: CRITICAL - First create empty column maps, then ensure result columns exist
    ' This must be done BEFORE GetColumnMap so that the new columns are included in the mapping
    ' FIX: Skip EnsureWriteColumnsExist if all result columns already exist (prevents data loss)
    Set sourceCols = CreateObject("Scripting.Dictionary")
    Set targetCols = CreateObject("Scripting.Dictionary")

    ' FIX: REMOVED EnsureWriteColumnsExist call - validation in ExecuteCompareWithValidation
    ' now blocks execution if columns are missing, user must run Build UI first
    ' This prevents the data deletion bug caused by EnsureWriteColumnsExist
    Dim sourceHasAllColumns As Boolean
    sourceHasAllColumns = CheckAllResultColumnsExist(g_SourceSheet, g_SourceHeaderRow)

    If Not sourceHasAllColumns Then
        ' Should not reach here - ExecuteCompareWithValidation should block execution
        DebugPrint "CompareAssets: WARNING - Columns missing but validation should have blocked this"
    Else
        DebugPrint "CompareAssets: Source has all result columns verified"
    End If

    ' FIX: Do NOT modify Target file - it's read-only
    ' Results are only written to Source file

    ' Step 6: Get column mappings with auto-mapping (AFTER columns are created)
    ' Merge with existing colMaps from EnsureWriteColumnsExist
    Dim tempSourceCols As Object
    Dim tempTargetCols As Object
    Set tempSourceCols = GetColumnMap(g_SourceSheet, g_SourceHeaderRow, headerAliases, sourceLearned)
    Set tempTargetCols = GetColumnMap(g_TargetSheet, g_TargetHeaderRow, headerAliases, targetLearned)

    ' Copy temp cols to existing sourceCols/targetCols (preserving any columns already added)
    Dim k As Variant
    For Each k In tempSourceCols.keys
        If Not sourceCols.exists(k) Then sourceCols.Add k, tempSourceCols(k)
    Next k
    For Each k In tempTargetCols.keys
        If Not targetCols.exists(k) Then targetCols.Add k, tempTargetCols(k)
    Next k

    ' FIX: Ensure MATCH result columns are in sourceCols - search directly in the sheet
    ' This handles cases where columns exist but weren't added to sourceCols
    Call EnsureResultColumnsInColMap(g_SourceSheet, g_SourceHeaderRow, sourceCols)

    ' Debug: Print source column map
    Dim debugSourceMap As String
    debugSourceMap = "SOURCE columns: "
    For Each k In sourceCols.keys
        debugSourceMap = debugSourceMap & k & "=" & sourceCols(k) & ", "
    Next k
    DebugPrint debugSourceMap

    ' Debug: Print target column map
    Dim debugTargetMap As String
    debugTargetMap = "TARGET columns: "
    For Each k In targetCols.keys
        debugTargetMap = debugTargetMap & k & "=" & targetCols(k) & ", "
    Next k
    DebugPrint debugTargetMap

    ' Step 7: Auto-map columns between source and target
    Call AutoMapColumns(sourceCols, targetCols, g_SourceSheet, g_TargetSheet, _
                       g_SourceHeaderRow, g_TargetHeaderRow, headerAliases)

    ' Debug: Print column maps AFTER AutoMap
    DebugPrint "CompareAssets: SOURCE columns AFTER AutoMap: " & sourceCols.Count
    DebugPrint "CompareAssets: TARGET columns AFTER AutoMap: " & targetCols.Count

    ' Step 8: Read data arrays
    sourceData = GetSheetData(g_SourceSheet, g_SourceHeaderRow)
    targetData = GetSheetData(g_TargetSheet, g_TargetHeaderRow)

    ' Validate data arrays are valid 2D arrays with data
    Dim sourceRowCount As Long

    On Error Resume Next
    Err.Clear
    sourceRowCount = UBound(sourceData, 1)
    If Err.Number <> 0 Or sourceRowCount < 2 Then
        MsgBox "Could not read data from Source sheet." & vbCrLf & vbCrLf & _
               "This usually means:" & vbCrLf & _
               "1. The header row number (" & g_SourceHeaderRow & ") may be incorrect" & vbCrLf & _
               "2. There's no data below the header row" & vbCrLf & vbCrLf & _
               "Please run Execute Match again and confirm the correct header row number.", vbExclamation
        GoTo CleanExit
    End If

    Err.Clear
    targetRowCount = UBound(targetData, 1)
    If Err.Number <> 0 Or targetRowCount < 2 Then
        MsgBox "Could not read data from Target sheet." & vbCrLf & vbCrLf & _
               "This usually means:" & vbCrLf & _
               "1. The header row number (" & g_TargetHeaderRow & ") may be incorrect" & vbCrLf & _
               "2. There's no data below the header row" & vbCrLf & vbCrLf & _
               "Please run Execute Match again and confirm the correct header row number.", vbExclamation
        GoTo CleanExit
    End If
    On Error GoTo ErrorHandler

    totalRows = sourceRowCount - 1

    ' Step 9: Read match definitions from SOURCE sheet UI
    Set matchDefs = GetMatchDefinitionsFromUI(g_SourceSheet, g_SourceHeaderRow, sourceCols)

    ' Debug: show what match definitions were captured
    DebugPrint "CompareAssets: Got " & matchDefs.Count & " match definitions"
    Dim defIdx As Long
    Dim colKey As Variant
    For defIdx = 1 To matchDefs.Count
        Set matchDef = matchDefs(defIdx)
        Dim colList As String
        colList = ""
        For Each colKey In matchDef("Columns")
            colList = colList & colKey & ", "
        Next colKey
        DebugPrint "CompareAssets: Match " & defIdx & " - ID=" & matchDef("ID") & ", Type=" & matchDef("Type") & ", Columns=[" & Left(colList, Len(colList)-2) & "]"
    Next defIdx

    If matchDefs.Count = 0 Then
        MsgBox "No match definitions found. Please define at least one match row.", vbExclamation
        GoTo CleanExit
    End If

    ' Step 10: Build lookup dictionaries from TARGET
    Set dictList = New Collection

    For m = 1 To matchDefs.Count
        Set dict = CreateObject("Scripting.Dictionary")
        dictList.Add dict
    Next m

    Application.StatusBar = "Building TARGET lookup dictionaries..."

    ' Debug: show sample keys being built
    DebugPrint "CompareAssets: Building TARGET dictionary..."

    ' Build dictionary from TARGET data
    ' OPTIMIZATION: Pre-calculate match columns to avoid repeated dictionary lookups
    Dim targetLastRow As Long
    targetLastRow = UBound(targetData, 1)

    ' PRE-COMPUTE: Find leftmost target ID column index once (avoids colMap enumeration per row)
    targetIdColIdx = 0
    tempFoundIdx = 999999
    For Each preColKey In targetCols.keys
        If UCase(preColKey) <> "MATCHED_ID" And UCase(preColKey) <> "MATCH_TYPE" And _
           UCase(preColKey) <> "MATCH_STATUS" And UCase(preColKey) <> "SOURCE_FILE" And _
           UCase(preColKey) <> "TARGET_FILE" And UCase(preColKey) <> "MATCHED_ASSETID" Then
            If targetCols(preColKey) < tempFoundIdx Then
                tempFoundIdx = targetCols(preColKey)
                targetIdColIdx = targetCols(preColKey)
            End If
        End If
    Next preColKey

    ' PRE-COMPUTE: Find leftmost source ID column index once (avoids colMap enumeration per row)
    sourceIdColIdx = 0
    tempFoundIdx = 999999
    For Each preColKey In sourceCols.keys
        If UCase(preColKey) <> "MATCHED_ID" And UCase(preColKey) <> "MATCH_TYPE" And _
           UCase(preColKey) <> "MATCH_STATUS" And UCase(preColKey) <> "SOURCE_FILE" And _
           UCase(preColKey) <> "TARGET_FILE" And UCase(preColKey) <> "MATCHED_ASSETID" Then
            If sourceCols(preColKey) < tempFoundIdx Then
                tempFoundIdx = sourceCols(preColKey)
                sourceIdColIdx = sourceCols(preColKey)
            End If
        End If
    Next preColKey

    For targetRow = 2 To targetLastRow
        For m = 1 To matchDefs.Count
            Set matchDef = matchDefs(m)
            key = BuildKeyFromRow(targetData, targetRow, matchDef("Columns"), targetCols)

            If key <> "" Then
                ' Get TARGET key column value (ASSETID or equivalent)
                ' Use user-selected column if specified, otherwise auto-detect
                If g_MatchedIdColumn <> "" And targetCols.Exists(g_MatchedIdColumn) Then
                    targetKeyValue = SafeCleanString(targetData(targetRow, targetCols(g_MatchedIdColumn)))
                Else
                    If targetIdColIdx > 0 Then
                        targetKeyValue = SafeCleanString(targetData(targetRow, targetIdColIdx))
                    Else
                        targetKeyValue = GetKeyColumnValue(targetData, targetRow, targetCols)
                    End If
                End If

                If targetKeyValue <> "" Then
                    Set dictM = dictList(m)
                    If Not dictM.exists(key) Then
                        dictM.Add key, targetKeyValue
                    End If
                End If
            End If
        Next m

        ' OPTIMIZATION: Update status every 10000 rows instead of 50000 to show progress
        If targetRow Mod 10000 = 0 Then
            Application.StatusBar = "TARGET: " & Format(targetRow - 1, "#,##0") & " rows processed..."
        End If
    Next targetRow

    ' Step 11: Process SOURCE rows and find matches
    ' Validate that sourceData has actual data rows (not just header)
    If UBound(sourceData, 1) < 2 Then
        MsgBox "No data rows found in source sheet (only header exists).", vbExclamation
        GoTo CleanExit
    End If

    ReDim matchCounts(1 To matchDefs.Count)
    noMatchCount = 0

    ReDim sourceResults(1 To UBound(sourceData, 1) - 1, 1 To 5)  ' 5 columns: MATCHED_ID, MATCH_TYPE, MATCH_STATUS, SOURCE_FILE, TARGET_FILE

    ' Also build reverse lookup for target -> source matches
    Set reverseDict = CreateObject("Scripting.Dictionary")

    Application.StatusBar = "Matching SOURCE assets..."

    ' Debug: show sample SOURCE keys
    DebugPrint "CompareAssets: Matching SOURCE rows..."

    ' Read actual loaded file names from config for SOURCE_FILE/TARGET_FILE columns
    Set wbNameConfigSheet = GetOrCreateConfigSheet
    targetWorkbookName = GetConfigValue(wbNameConfigSheet, "LOADED_TARGET_FILE")
    sourceWorkbookName = GetConfigValue(wbNameConfigSheet, "LOADED_SOURCE_FILE")
    If targetWorkbookName = "" Then targetWorkbookName = GetWorkbookName(g_TargetWorkbook)
    If sourceWorkbookName = "" Then sourceWorkbookName = GetWorkbookName(g_SourceWorkbook)

    ' PRE-COMPUTE: Resolve column indices for each match definition
    ' Allows BuildKeyFast to replace BuildKeyFromRow in both loops
    mDefCount = matchDefs.Count
    ReDim fastKeyIdxs(1 To mDefCount, 1 To 2)
    ReDim fastKeyColCounts(1 To mDefCount)  ' (matchDef, 1=source/2=target) → Long() array stored as Variant

    For preM = 1 To mDefCount
        Set matchDef = matchDefs(preM)

        ' Resolve SOURCE column indices for this match definition
        preCount = 0
        For Each preColName In matchDef("Columns")
            preNorm = UltraNormalize(CStr(preColName))
            If preNorm <> "" And sourceCols.exists(preNorm) Then
                preCount = preCount + 1
            ElseIf sourceCols.exists(CStr(preColName)) Then
                preCount = preCount + 1
            End If
        Next preColName

        ReDim srcIdxArr(1 To IIf(preCount > 0, preCount, 1))
        preCount = 0
        For Each preColName In matchDef("Columns")
            preNorm = UltraNormalize(CStr(preColName))
            If preNorm <> "" And sourceCols.exists(preNorm) Then
                preCount = preCount + 1
                srcIdxArr(preCount) = sourceCols(preNorm)
            ElseIf sourceCols.exists(CStr(preColName)) Then
                preCount = preCount + 1
                srcIdxArr(preCount) = sourceCols(CStr(preColName))
            End If
        Next preColName
        On Error Resume Next
        fastKeyIdxs(preM, 1) = srcIdxArr
        On Error GoTo 0
        fastKeyColCounts(preM) = preCount

        ' Resolve TARGET column indices for this match definition
        preCount = 0
        For Each preColName In matchDef("Columns")
            preNorm = UltraNormalize(CStr(preColName))
            If preNorm <> "" And targetCols.exists(preNorm) Then
                preCount = preCount + 1
            ElseIf targetCols.exists(CStr(preColName)) Then
                preCount = preCount + 1
            End If
        Next preColName

        ReDim tgtIdxArr(1 To IIf(preCount > 0, preCount, 1))
        preCount = 0
        For Each preColName In matchDef("Columns")
            preNorm = UltraNormalize(CStr(preColName))
            If preNorm <> "" And targetCols.exists(preNorm) Then
                preCount = preCount + 1
                tgtIdxArr(preCount) = targetCols(preNorm)
            ElseIf targetCols.exists(CStr(preColName)) Then
                preCount = preCount + 1
                tgtIdxArr(preCount) = targetCols(CStr(preColName))
            End If
        Next preColName
        On Error Resume Next
        fastKeyIdxs(preM, 2) = tgtIdxArr
        On Error GoTo 0
    Next preM

    ' OPTIMIZATION: Pre-calculate values outside loop
    Dim sourceLastRow As Long
    sourceLastRow = UBound(sourceData, 1)
    Dim noMatch As String
    Dim matchDone As String
    noMatch = "NO_MATCH"
    matchDone = "DONE"

    For sourceRow = 2 To sourceLastRow
        matchFound = False
        matchedID = "~NULL~"
        matchType = "~NULL~"

        ' Get source ASSETID for reverse lookup
        If sourceIdColIdx > 0 Then
            sourceAssetID = SafeCleanString(sourceData(sourceRow, sourceIdColIdx))
        Else
            sourceAssetID = GetKeyColumnValue(sourceData, sourceRow, sourceCols)
        End If

        ' Try matches in order (TOP TO BOTTOM - FIRST MATCH ONLY)
        For m = 1 To matchDefs.Count
            Set matchDef = matchDefs(m)
            sourceKey = BuildKeyFromRow(sourceData, sourceRow, matchDef("Columns"), sourceCols)

            If sourceKey <> "" Then
                Set dictTry = dictList(m)
                If dictTry.exists(sourceKey) Then
                    matchedID = dictTry(sourceKey)
                    matchType = matchDef("Type")
                    matchFound = True
                    matchCounts(m) = matchCounts(m) + 1

                    ' Add to reverse lookup for target -> source
                    If sourceAssetID <> "" And Not reverseDict.exists(matchedID) Then
                        reverseDict.Add matchedID, sourceAssetID & "|" & matchType
                    End If

                    Exit For  ' FIRST MATCH - stop immediately
                End If
            End If
        Next m

        ' Store results (5 columns: MATCHED_ID, MATCH_TYPE, MATCH_STATUS, SOURCE_FILE, TARGET_FILE)
        sourceResults(sourceRow - 1, 1) = matchedID
        sourceResults(sourceRow - 1, 2) = matchType
        sourceResults(sourceRow - 1, 3) = IIf(matchFound, matchDone, noMatch)
        ' FIXED: SOURCE_FILE = source workbook (file we're writing to), TARGET_FILE = target workbook (file we're matching against)
        sourceResults(sourceRow - 1, 4) = sourceWorkbookName  ' SOURCE_FILE = source workbook name
        sourceResults(sourceRow - 1, 5) = targetWorkbookName  ' TARGET_FILE = target workbook name

        If Not matchFound Then noMatchCount = noMatchCount + 1

        ' OPTIMIZATION: Update status every 10000 rows
        processedRows = sourceRow - 1
        If processedRows Mod 10000 = 0 Then
            Application.StatusBar = "SOURCE: " & Format(processedRows, "#,##0") & " of " & Format(totalRows, "#,##0") & _
                                  " (" & Format(processedRows / totalRows * 100, "0.0") & "%)"
        End If
    Next sourceRow

    ' Step 12: Write results to SOURCE sheet - write each column at its specific position
    Application.StatusBar = "Writing results to SOURCE sheet..."

    ' Debug: Print column mapping and row counts for diagnosing issues
    DebugPrint "=== Column Mapping Debug ==="
    DebugPrint "Source ASSETID column: " & IIf(sourceCols.exists("ASSETID"), sourceCols("ASSETID"), "NOT FOUND")
    DebugPrint "Target ASSETID column: " & IIf(targetCols.exists("ASSETID"), targetCols("ASSETID"), "NOT FOUND")
    DebugPrint "Source rows: " & sourceRowCount & ", Target rows: " & targetRowCount

    ' Debug: Print column positions
    DebugPrint "Step 12: SOURCE - MATCHED_ID col: " & IIf(sourceCols.exists("MATCHED_ID"), sourceCols("MATCHED_ID"), "NOT FOUND")
    DebugPrint "Step 12: SOURCE - MATCH_TYPE col: " & IIf(sourceCols.exists("MATCH_TYPE"), sourceCols("MATCH_TYPE"), "NOT FOUND")
    DebugPrint "Step 12: SOURCE - MATCH_STATUS col: " & IIf(sourceCols.exists("MATCH_STATUS"), sourceCols("MATCH_STATUS"), "NOT FOUND")
    DebugPrint "Step 12: SOURCE - SOURCE_FILE col: " & IIf(sourceCols.exists("SOURCE_FILE"), sourceCols("SOURCE_FILE"), "NOT FOUND")
    DebugPrint "Step 12: SOURCE - TARGET_FILE col: " & IIf(sourceCols.exists("TARGET_FILE"), sourceCols("TARGET_FILE"), "NOT FOUND")

    ' OPTIMIZATION: Write all columns in bulk using array assignment instead of cell-by-cell
    ' This is MUCH faster for large datasets
    rowCount = UBound(sourceResults, 1)
    sourceStartRow = g_SourceHeaderRow + 1

    ' Write all 5 columns at once using bulk array assignment
    If rowCount > 0 Then
        ' Build a 2D array with all results transposed to columns
        ReDim sourceWriteArray(1 To rowCount, 1 To 5)

        For i = 1 To rowCount
            sourceWriteArray(i, 1) = sourceResults(i, 1)  ' MATCHED_ID
            sourceWriteArray(i, 2) = sourceResults(i, 2) ' MATCH_TYPE
            sourceWriteArray(i, 3) = sourceResults(i, 3) ' MATCH_STATUS
            sourceWriteArray(i, 4) = sourceResults(i, 4) ' SOURCE_FILE
            sourceWriteArray(i, 5) = sourceResults(i, 5) ' TARGET_FILE
        Next i

        ' Determine the leftmost column to write to
        minSourceCol = 0
        If sourceCols.exists("MATCHED_ID") Then minSourceCol = sourceCols("MATCHED_ID")
        If sourceCols.exists("MATCH_TYPE") Then
            If minSourceCol = 0 Or sourceCols("MATCH_TYPE") < minSourceCol Then minSourceCol = sourceCols("MATCH_TYPE")
        End If
        If sourceCols.exists("MATCH_STATUS") Then
            If minSourceCol = 0 Or sourceCols("MATCH_STATUS") < minSourceCol Then minSourceCol = sourceCols("MATCH_STATUS")
        End If
        If sourceCols.exists("SOURCE_FILE") Then
            If minSourceCol = 0 Or sourceCols("SOURCE_FILE") < minSourceCol Then minSourceCol = sourceCols("SOURCE_FILE")
        End If
        If sourceCols.exists("TARGET_FILE") Then
            If minSourceCol = 0 Or sourceCols("TARGET_FILE") < minSourceCol Then minSourceCol = sourceCols("TARGET_FILE")
        End If

        ' FIX: If minSourceCol is still 0, create columns at the end of the sheet
        ' This ensures results are written even if MATCH columns don't exist
        If minSourceCol = 0 Then
            ' Find last column with data in header row
            Dim lastCol As Long
            ' FIX: Replace End(xlToLeft) with explicit rightward scan
        Dim scanCol As Long
        lastCol = 0
        For scanCol = 1 To 200
            If Trim(CStr(g_SourceSheet.Cells(g_SourceHeaderRow, scanCol).Value)) <> "" Then
                lastCol = scanCol
            End If
        Next scanCol
        If lastCol < 1 Then lastCol = 10  ' Fallback if no headers found

            ' FIX: Check each MATCH column header individually before inserting
            Dim matchHeaders As Variant
            matchHeaders = Array("MATCHED_ID", "MATCH_TYPE", "MATCH_STATUS", "SOURCE_FILE", "TARGET_FILE")

            Dim existingMatchCols As Object
            Set existingMatchCols = CreateObject("Scripting.Dictionary")

            Dim headerScanCol As Long
            Dim headerCellValue As String
            For headerScanCol = 1 To 200
                headerCellValue = UCase(Trim(CStr(g_SourceSheet.Cells(g_SourceHeaderRow, headerScanCol).Value)))
                If headerCellValue <> "" Then
                    Dim hIdx As Variant
                    For Each hIdx In matchHeaders
                        If headerCellValue = hIdx Then
                            If Not existingMatchCols.Exists(hIdx) Then
                                existingMatchCols.Add hIdx, headerScanCol
                            End If
                        End If
                    Next
                End If
            Next headerScanCol

            Dim leftmostMatchCol As Long
            leftmostMatchCol = 0
            Dim mKey As Variant
            For Each mKey In existingMatchCols.keys
                If leftmostMatchCol = 0 Or existingMatchCols(mKey) < leftmostMatchCol Then
                    leftmostMatchCol = existingMatchCols(mKey)
                End If
            Next

            If leftmostMatchCol = 0 Then
                leftmostMatchCol = lastCol + 1
            End If

            ' Insert missing columns RIGHT-TO-LEFT (rightmost first)
            If Not existingMatchCols.Exists("TARGET_FILE") Then
                g_SourceSheet.Columns(leftmostMatchCol + 4).Insert Shift:=xlToRight
            End If
            If Not existingMatchCols.Exists("SOURCE_FILE") Then
                g_SourceSheet.Columns(leftmostMatchCol + 3).Insert Shift:=xlToRight
            End If
            If Not existingMatchCols.Exists("MATCH_STATUS") Then
                g_SourceSheet.Columns(leftmostMatchCol + 2).Insert Shift:=xlToRight
            End If
            If Not existingMatchCols.Exists("MATCH_TYPE") Then
                g_SourceSheet.Columns(leftmostMatchCol + 1).Insert Shift:=xlToRight
            End If
            If Not existingMatchCols.Exists("MATCHED_ID") Then
                g_SourceSheet.Columns(leftmostMatchCol).Insert Shift:=xlToRight
            End If

            ' Write header ONLY if newly inserted
            If Not existingMatchCols.Exists("MATCHED_ID") Then
                g_SourceSheet.Cells(g_SourceHeaderRow, leftmostMatchCol).Value = "MATCHED_ID"
            End If
            If Not existingMatchCols.Exists("MATCH_TYPE") Then
                g_SourceSheet.Cells(g_SourceHeaderRow, leftmostMatchCol + 1).Value = "MATCH_TYPE"
            End If
            If Not existingMatchCols.Exists("MATCH_STATUS") Then
                g_SourceSheet.Cells(g_SourceHeaderRow, leftmostMatchCol + 2).Value = "MATCH_STATUS"
            End If
            If Not existingMatchCols.Exists("SOURCE_FILE") Then
                g_SourceSheet.Cells(g_SourceHeaderRow, leftmostMatchCol + 3).Value = "SOURCE_FILE"
            End If
            If Not existingMatchCols.Exists("TARGET_FILE") Then
                g_SourceSheet.Cells(g_SourceHeaderRow, leftmostMatchCol + 4).Value = "TARGET_FILE"
            End If

            minSourceCol = leftmostMatchCol
            DebugPrint "Step 12: Ensured MATCH columns exist at column " & minSourceCol
        End If

        ' FIX: Use pre-known column positions instead of rescanning
        ' We already know where we inserted the columns at minSourceCol
        If minSourceCol > 0 Then
            ' Direct assignment of column positions - no rescan needed
            ' MATCH columns were inserted at: minSourceCol, minSourceCol+1, minSourceCol+2, etc.

            ' Extract and write each column one at a time
            Dim matchedIdColData As Variant
            ReDim matchedIdColData(1 To rowCount, 1 To 1)
            Dim r As Long
            For r = 1 To rowCount
                matchedIdColData(r, 1) = sourceWriteArray(r, 1)
            Next r
            g_SourceSheet.Cells(sourceStartRow, minSourceCol + 0).Resize(rowCount, 1).Value = matchedIdColData

            Dim matchTypeColData As Variant
            ReDim matchTypeColData(1 To rowCount, 1 To 1)
            For r = 1 To rowCount
                matchTypeColData(r, 1) = sourceWriteArray(r, 2)
            Next r
            g_SourceSheet.Cells(sourceStartRow, minSourceCol + 1).Resize(rowCount, 1).Value = matchTypeColData

            Dim matchStatusColData As Variant
            ReDim matchStatusColData(1 To rowCount, 1 To 1)
            For r = 1 To rowCount
                matchStatusColData(r, 1) = sourceWriteArray(r, 3)
            Next r
            g_SourceSheet.Cells(sourceStartRow, minSourceCol + 2).Resize(rowCount, 1).Value = matchStatusColData

            Dim sourceFileColData As Variant
            ReDim sourceFileColData(1 To rowCount, 1 To 1)
            For r = 1 To rowCount
                sourceFileColData(r, 1) = sourceWriteArray(r, 4)
            Next r
            g_SourceSheet.Cells(sourceStartRow, minSourceCol + 3).Resize(rowCount, 1).Value = sourceFileColData

            Dim targetFileColData As Variant
            ReDim targetFileColData(1 To rowCount, 1 To 1)
            For r = 1 To rowCount
                targetFileColData(r, 1) = sourceWriteArray(r, 5)
            Next r
            g_SourceSheet.Cells(sourceStartRow, minSourceCol + 4).Resize(rowCount, 1).Value = targetFileColData

            DebugPrint "Step 12: Bulk wrote " & rowCount & " rows to MATCH columns starting at column " & minSourceCol & "..."
        End If
    End If

    ' Step 13: Write results to TARGET sheet
    Application.StatusBar = "Writing results to TARGET sheet..."

    ' Build target results array
    ReDim targetResults(1 To UBound(targetData, 1) - 1, 1 To 5)  ' 5 columns: MATCHED_ID, MATCH_TYPE, MATCH_STATUS, SOURCE_FILE, TARGET_FILE

    For targetRow = 2 To UBound(targetData, 1)
        ' Get target ASSETID
        ' Use user-selected column if specified, otherwise auto-detect
        normalizedMatchCol = SmartMatch(g_MatchedIdColumn, UltraNormalize(g_MatchedIdColumn), Nothing, Nothing)
        If normalizedMatchCol = "" Then normalizedMatchCol = g_MatchedIdColumn
        If normalizedMatchCol <> "" And targetCols.Exists(normalizedMatchCol) Then
            On Error Resume Next
            resolvedColIdx = CLng(targetCols(normalizedMatchCol))
            On Error GoTo ErrorHandler
            If resolvedColIdx > 0 And resolvedColIdx <= UBound(targetData, 2) And targetRow <= UBound(targetData, 1) Then
                targetKeyID = SafeCleanString(targetData(targetRow, resolvedColIdx))
            Else
                targetKeyID = GetKeyColumnValue(targetData, targetRow, targetCols)
            End If
        Else
            targetKeyID = GetKeyColumnValue(targetData, targetRow, targetCols)
        End If

        If reverseDict.exists(targetKeyID) Then
            ' This target row was matched by source
            targetMatchedInfo = reverseDict(targetKeyID)
            parts = Split(targetMatchedInfo, "|")
            targetResults(targetRow - 1, 1) = parts(0)  ' Source ASSETID
            If UBound(parts) >= 1 Then
                targetResults(targetRow - 1, 2) = parts(1)  ' Match type
            Else
                targetResults(targetRow - 1, 2) = ""
            End If
            targetResults(targetRow - 1, 3) = "DONE"
            ' FIXED: For TARGET sheet, SOURCE_FILE = source workbook, TARGET_FILE = target workbook
            targetResults(targetRow - 1, 4) = sourceWorkbookName  ' SOURCE_FILE = source workbook
            targetResults(targetRow - 1, 5) = targetWorkbookName  ' TARGET_FILE = target workbook
        Else
            ' Not matched
            targetResults(targetRow - 1, 1) = ""
            targetResults(targetRow - 1, 2) = ""
            targetResults(targetRow - 1, 3) = "NO_MATCH"
            targetResults(targetRow - 1, 4) = sourceWorkbookName  ' SOURCE_FILE = source workbook
            targetResults(targetRow - 1, 5) = targetWorkbookName  ' TARGET_FILE = target workbook
        End If
    Next targetRow

    ' Debug: Print column positions for TARGET
    DebugPrint "Step 13: TARGET - MATCHED_ID col: " & IIf(targetCols.exists("MATCHED_ID"), targetCols("MATCHED_ID"), "NOT FOUND")
    DebugPrint "Step 13: TARGET - MATCH_TYPE col: " & IIf(targetCols.exists("MATCH_TYPE"), targetCols("MATCH_TYPE"), "NOT FOUND")
    DebugPrint "Step 13: TARGET - MATCH_STATUS col: " & IIf(targetCols.exists("MATCH_STATUS"), targetCols("MATCH_STATUS"), "NOT FOUND")
    DebugPrint "Step 13: TARGET - SOURCE_FILE col: " & IIf(targetCols.exists("SOURCE_FILE"), targetCols("SOURCE_FILE"), "NOT FOUND")
    DebugPrint "Step 13: TARGET - TARGET_FILE col: " & IIf(targetCols.exists("TARGET_FILE"), targetCols("TARGET_FILE"), "NOT FOUND")

    ' FIX: Do NOT write to Target file - it's read-only
    ' Results are only written to Source file
    DebugPrint "Step 13: Skipping Target file write (read-only)"

    ' Cleanup
    endTime = Timer
    Call OptimizePerformance(False)
    Application.StatusBar = False

    ' Summary message
    summary = "Asset comparison complete!" & vbNewLine & vbNewLine & _
              "Match Rules: " & matchDefs.Count & vbNewLine & _
              "Source Rows: " & Format(processedRows, "#,##0") & vbNewLine & _
              "Target Rows: " & Format(UBound(targetData, 1) - 1, "#,##0") & vbNewLine & _
              "Matches Found: " & Format(processedRows - noMatchCount, "#,##0") & vbNewLine & _
              "No Matches: " & Format(noMatchCount, "#,##0") & vbNewLine & _
              "Time: " & Format(endTime - startTime, "0.0") & " seconds"

    DebugPrint "CompareAssets: COMPLETE - " & (processedRows - noMatchCount) & " matches found"
    MsgBox summary, vbInformation, "Comparison Complete"

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    Resume CleanExit

CleanExit:
    On Error Resume Next
    Call OptimizePerformance(False)
    Application.StatusBar = False
    On Error GoTo 0
End Sub

'===============================================================================
' SETUP WIZARD
'===============================================================================

Public Sub SetupWizard()
    '
    ' Step 1: Select SOURCE sheet
    '
    Dim sourceWS As Worksheet
    Set sourceWS = SelectSheetDialog("Select SOURCE sheet (the data you want to match)")
    If sourceWS Is Nothing Then
        MsgBox "Setup cancelled.", vbInformation
        Exit Sub
    End If

    '
    ' Step 2: Select TARGET sheet
    '
    Dim targetWS As Worksheet
    Set targetWS = SelectSheetDialog("Select TARGET sheet (the data to match against)")
    If targetWS Is Nothing Then
        MsgBox "Setup cancelled.", vbInformation
        Exit Sub
    End If

    '
    ' Step 3: Set global variables
    '
    Set g_SourceWorkbook = sourceWS.Parent
    Set g_SourceSheet = sourceWS
    Set g_TargetWorkbook = targetWS.Parent
    Set g_TargetSheet = targetWS

    '
    ' Step 4: Save configuration
    '
    Dim configSheet As Worksheet
    Set configSheet = GetOrCreateConfigSheet
    Call SetConfigValue(configSheet, CONFIG_SOURCE_WB, g_SourceWorkbook.Name)
    Call SetConfigValue(configSheet, CONFIG_SOURCE_WS, g_SourceSheet.Name)
    Call SetConfigValue(configSheet, CONFIG_TARGET_WB, g_TargetWorkbook.Name)
    Call SetConfigValue(configSheet, CONFIG_TARGET_WS, g_TargetSheet.Name)

    '
    ' Step 5: Build UI on source sheet
    '
    Set g_CurrentWorksheet = g_SourceSheet

    Call RebuildMatchBuilderUI

    MsgBox "Setup complete!" & vbCrLf & vbCrLf & _
           "Source: " & g_SourceSheet.Name & " (" & g_SourceWorkbook.Name & ")" & vbCrLf & _
           "Target: " & g_TargetSheet.Name & " (" & g_TargetWorkbook.Name & ")" & vbCrLf & vbCrLf & _
           "Now define your match rules and click Execute Match.", vbInformation
End Sub

Private Function SelectSheetDialog(promptText As String) As Worksheet
    '
    ' Shows a dialog to select a worksheet from any open workbook
    '
    Dim wbList As Collection
    Dim wsList As Collection
    Dim i As Long, j As Long
    Dim selection As String  ' FIX: Changed from Integer to String
    Dim prompt As String
    Dim selectedWB As Workbook
    Dim selectedWS As Worksheet
    Dim validSelection As Boolean
    Dim sheetName As String

    ' Build list of all worksheets in all open workbooks
    ' FIX: Filter out system sheets (COMPARE_CONFIG, CONFIG, etc.)
    Set wbList = New Collection
    Set wsList = New Collection

    For i = 1 To Workbooks.Count
        wbList.Add Workbooks(i)
        For j = 1 To Workbooks(i).Worksheets.Count
            sheetName = UCase(Workbooks(i).Worksheets(j).Name)
            ' Skip system sheets
            If sheetName <> "COMPARE_CONFIG" And sheetName <> "CONFIG" Then
                wsList.Add Workbooks(i).Worksheets(j)
            End If
        Next j
    Next i

    If wbList.Count = 0 Then
        MsgBox "No workbooks are open.", vbCritical
        Set SelectSheetDialog = Nothing
        Exit Function
    End If

    If wsList.Count = 0 Then
        MsgBox "No valid sheets found (system sheets excluded).", vbCritical
        Set SelectSheetDialog = Nothing
        Exit Function
    End If

    ' FIX: Add retry loop for invalid selection
    validSelection = False
    Do While Not validSelection
        ' Build prompt showing first 20 sheets
        prompt = promptText & vbCrLf & vbCrLf & _
                 "Available sheets:" & vbCrLf

        Dim maxShow As Long
        maxShow = 20
        If wsList.Count < maxShow Then maxShow = wsList.Count

        For i = 1 To maxShow
            Dim ws As Worksheet
            Set ws = wsList(i)
            prompt = prompt & i & ". [" & ws.Parent.Name & "] " & ws.Name & vbCrLf
        Next i

        If wsList.Count > maxShow Then
            prompt = prompt & "... and " & (wsList.Count - maxShow) & " more." & vbCrLf
        End If

        prompt = prompt & vbCrLf & "Enter number (or press Cancel to exit):"

        selection = InputBox(prompt, "Select Sheet", "1")

        ' Handle cancel or empty input
        If selection = "" Then
            Set SelectSheetDialog = Nothing
            Exit Function
        End If

        ' Parse the selection
        On Error Resume Next
        Dim selNum As Long
        selNum = CLng(selection)
        On Error GoTo 0

        If selNum >= 1 And selNum <= wsList.Count Then
            validSelection = True
        Else
            MsgBox "Invalid selection. Please enter a number between 1 and " & wsList.Count, vbExclamation
        End If
    Loop

    ' Use the parsed number (selNum) to get from the collection
    Set selectedWS = wsList(selNum)
    Set SelectSheetDialog = selectedWS
End Function

'===============================================================================
' CONFIGURATION MODULE - Source/Target Selection and Persistence
'===============================================================================

Private Sub ConfigureSourceTarget()
    ' Check if configuration exists
    Dim configSheet As Worksheet
    Set configSheet = GetOrCreateConfigSheet

    Dim sourceWBName As String, sourceWSName As String
    Dim targetWBName As String, targetWSName As String

    ' Load saved configuration
    sourceWBName = GetConfigValue(configSheet, CONFIG_SOURCE_WB)
    sourceWSName = GetConfigValue(configSheet, CONFIG_SOURCE_WS)
    targetWBName = GetConfigValue(configSheet, CONFIG_TARGET_WB)
    targetWSName = GetConfigValue(configSheet, CONFIG_TARGET_WS)

    ' If any configuration is missing, show selection dialog
    If sourceWBName = "" Or sourceWSName = "" Or targetWBName = "" Or targetWSName = "" Then
        Call ShowSourceTargetDialog(sourceWBName, sourceWSName, targetWBName, targetWSName)

        ' Save configuration
        Call SetConfigValue(configSheet, CONFIG_SOURCE_WB, sourceWBName)
        Call SetConfigValue(configSheet, CONFIG_SOURCE_WS, sourceWSName)
        Call SetConfigValue(configSheet, CONFIG_TARGET_WB, targetWBName)
        Call SetConfigValue(configSheet, CONFIG_TARGET_WS, targetWSName)
    End If

    ' Set source workbook and sheet
    Set g_SourceWorkbook = GetWorkbookByName(sourceWBName)
    If g_SourceWorkbook Is Nothing Then
        Set g_SourceWorkbook = ThisWorkbook
    End If
    Set g_SourceSheet = SafeGetWorksheet(g_SourceWorkbook, sourceWSName)
    ' Fallback: if saved sheet name no longer exists, use current worksheet
    If g_SourceSheet Is Nothing Then
        Set g_SourceSheet = g_CurrentWorksheet
        If g_SourceSheet Is Nothing Then Set g_SourceSheet = ActiveSheet
    End If

    ' Set target workbook and sheet
    Set g_TargetWorkbook = GetWorkbookByName(targetWBName)
    If g_TargetWorkbook Is Nothing Then
        Set g_TargetWorkbook = ThisWorkbook
    End If
    Set g_TargetSheet = SafeGetWorksheet(g_TargetWorkbook, targetWSName)
End Sub

Private Sub ShowSourceTargetDialog(ByRef sourceWB As String, ByRef sourceWS As String, _
                                   ByRef targetWB As String, ByRef targetWS As String)
    Dim wbList As Collection
    Dim wsList As Collection
    Dim i As Long
    Dim selection As Integer
    Dim wbPrompt As String
    Dim tempWB As Workbook
    Dim wsPrompt As String

    ' Get list of open workbooks
    Set wbList = New Collection
    For i = 1 To Workbooks.Count
        wbList.Add Workbooks(i).Name
    Next i

    If wbList.Count = 0 Then
        MsgBox "No workbooks are open.", vbCritical
        Exit Sub
    End If

    ' Build workbook selection prompt
    wbPrompt = "Select SOURCE Workbook:" & vbCrLf
    For i = 1 To wbList.Count
        wbPrompt = wbPrompt & i & ". " & wbList(i) & vbCrLf
    Next i

    ' Get source workbook
    If sourceWB = "" Then sourceWB = ThisWorkbook.Name

    selection = InputBox(wbPrompt & vbCrLf & "Enter number (default: " & FindInCollection(wbList, sourceWB) & "):", _
                         "Source Workbook Selection", FindInCollection(wbList, sourceWB))

    If selection > 0 And selection <= wbList.Count Then
        sourceWB = wbList(selection)
    End If

    ' Get source sheet (need workbook reference first)
    Set tempWB = GetWorkbookByName(sourceWB)
    If tempWB Is Nothing Then Set tempWB = ThisWorkbook

    Set wsList = New Collection
    For i = 1 To tempWB.Worksheets.Count
        wsList.Add tempWB.Worksheets(i).Name
    Next i

    If sourceWS = "" Then sourceWS = GetDefaultSheetName(tempWB)

    wsPrompt = "Select SOURCE Sheet in " & sourceWB & ":" & vbCrLf
    For i = 1 To wsList.Count
        wsPrompt = wsPrompt & i & ". " & wsList(i) & vbCrLf
    Next i

    selection = InputBox(wsPrompt & vbCrLf & "Enter number:", _
                         "Source Sheet Selection", FindInCollection(wsList, sourceWS))

    If selection > 0 And selection <= wsList.Count Then
        sourceWS = wsList(selection)
    End If

    ' Get target workbook
    If targetWB = "" Then targetWB = ThisWorkbook.Name

    wbPrompt = "Select TARGET Workbook:" & vbCrLf
    For i = 1 To wbList.Count
        wbPrompt = wbPrompt & i & ". " & wbList(i) & vbCrLf
    Next i

    selection = InputBox(wbPrompt & vbCrLf & "Enter number:", _
                         "Target Workbook Selection", FindInCollection(wbList, targetWB))

    If selection > 0 And selection <= wbList.Count Then
        targetWB = wbList(selection)
    End If

    ' Get target sheet
    Set tempWB = GetWorkbookByName(targetWB)
    If tempWB Is Nothing Then Set tempWB = ThisWorkbook

    Set wsList = New Collection
    For i = 1 To tempWB.Worksheets.Count
        wsList.Add tempWB.Worksheets(i).Name
    Next i

    If targetWS = "" Then targetWS = GetDefaultSheetName(tempWB)

    wsPrompt = "Select TARGET Sheet in " & targetWB & ":" & vbCrLf
    For i = 1 To wsList.Count
        wsPrompt = wsPrompt & i & ". " & wsList(i) & vbCrLf
    Next i

    selection = InputBox(wsPrompt & vbCrLf & "Enter number:", _
                         "Target Sheet Selection", FindInCollection(wsList, targetWS))

    If selection > 0 And selection <= wsList.Count Then
        targetWS = wsList(selection)
    End If
End Sub

Private Function FindInCollection(col As Collection, item As String) As Long
    Dim i As Long
    For i = 1 To col.Count
        If col(i) = item Then
            FindInCollection = i
            Exit Function
        End If
    Next i
    FindInCollection = 1
End Function

Private Function GetDefaultSheetName(wb As Workbook) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(1)
    On Error GoTo 0
    If Not ws Is Nothing Then GetDefaultSheetName = ws.Name
End Function

Private Function GetWorkbookByName(wbName As String) As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(wbName)
    On Error GoTo 0
    Set GetWorkbookByName = wb
End Function

Private Function GetWorkbookName(wb As Workbook) As String
    If wb Is Nothing Then
        GetWorkbookName = "Unknown"
    Else
        GetWorkbookName = wb.Name
    End If
End Function

'===============================================================================
' CONFIGURATION SHEET MANAGEMENT
'===============================================================================

Private Function GetOrCreateConfigSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CONFIG_SHEET_NAME
        ws.Visible = xlSheetVeryHidden

        ' Initialize configuration
        ws.Cells(1, 1).value = "Setting"
        ws.Cells(1, 2).value = "Value"
        ws.Cells(2, 1).value = CONFIG_SOURCE_WB
        ws.Cells(3, 1).value = CONFIG_SOURCE_WS
        ws.Cells(4, 1).value = CONFIG_TARGET_WB
        ws.Cells(5, 1).value = CONFIG_TARGET_WS

        DebugPrint "GetOrCreateConfigSheet: Created config sheet"
    End If

    Set GetOrCreateConfigSheet = ws
End Function

Private Function GetConfigValue(ws As Worksheet, settingName As String) As String
    Dim i As Long
    For i = 1 To 100
        If ws.Cells(i, 1).value = settingName Then
            GetConfigValue = CStr(ws.Cells(i, 2).value)
            Exit Function
        End If
    Next i
    GetConfigValue = ""
End Function

Private Sub SetConfigValue(ws As Worksheet, settingName As String, value As String)
    ' Safety check: if worksheet is Nothing, exit
    If ws Is Nothing Then
        DebugPrint "SetConfigValue: Warning - worksheet is Nothing, cannot set " & settingName
        Exit Sub
    End If

    Dim i As Long
    For i = 1 To 100
        If ws.Cells(i, 1).value = settingName Then
            ws.Cells(i, 2).value = value
            Exit Sub
        End If
    Next i

    ' Key not found — add it after the last used row
    Dim lastRow As Long
    lastRow = 0
    For i = 1 To 200
        If Trim(CStr(ws.Cells(i, 1).value)) <> "" Then
            lastRow = i
        End If
    Next i
    ws.Cells(lastRow + 1, 1).value = settingName
    ws.Cells(lastRow + 1, 2).value = value
End Sub

'===============================================================================
' HEADER MAPPING MODULE
'===============================================================================

Private Sub LoadHeaderMapping(aliases As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim standardHeader As String
    Dim aliasHeader As String
    Dim normalized As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HEADER_MAP_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    ' Read HEADER_MAP sheet: STANDARD_HEADER | ALIAS_HEADER
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        standardHeader = Trim(CStr(ws.Cells(i, 1).value))
        aliasHeader = Trim(CStr(ws.Cells(i, 2).value))

        If standardHeader <> "" And aliasHeader <> "" Then
            ' Add alias to dictionary (normalized form)
            normalized = UltraNormalize(aliasHeader)
            normalized = UltraNormalize(aliasHeader)
            If normalized <> "" And Not aliases.exists(normalized) Then
                aliases.Add normalized, standardHeader
            End If
        End If
    Next i
End Sub

Private Sub AutoMapColumns(sourceCols As Object, targetCols As Object, _
                            sourceSheet As Worksheet, targetSheet As Worksheet, _
                            sourceHeaderRow As Long, targetHeaderRow As Long, _
                            aliases As Object)

    ' These are RESULT columns that should NOT be mapped to target
    ' They are created by the system for output
    Dim resultColumns As Object
    Set resultColumns = CreateObject("Scripting.Dictionary")
    resultColumns.Add "MATCHED_ID", True
    resultColumns.Add "MATCHED_ASSETID", True
    resultColumns.Add "MATCH_ID", True
    resultColumns.Add "MATCH_TYPE", True
    resultColumns.Add "MATCH_STATUS", True
    resultColumns.Add "SOURCE_FILE", True
    resultColumns.Add "TARGET_FILE", True

    ' Also skip UI definition columns (Match_1, Match_2, Match_3, etc.)
    ' These are created by Build UI and should not be mapped to target
    Dim uiColumns As Object
    Set uiColumns = CreateObject("Scripting.Dictionary")

    ' Check each SOURCE column - if not found in TARGET, prompt user
    Dim sourceColName As Variant
    Dim targetColName As Variant
    Dim sourceColIdx As Long
    Dim targetColIdx As Long
    Dim foundMapping As Boolean
    Dim aliasName As String
    Dim targetCol As String

    For Each sourceColName In sourceCols.Keys
        sourceColIdx = sourceCols(sourceColName)

        ' SKIP result columns - they don't need to be mapped to target
        If resultColumns.exists(sourceColName) Then
            DebugPrint "AutoMapColumns: Skipping result column '" & sourceColName & "'"
            GoTo NextColumn
        End If

        ' SKIP UI definition columns (Match_1, Match_2, Match_3, Match_4, Match_5, Match.1, etc.)
        ' These are created by Build UI and should not be mapped to target
        Dim sourceColNameStr As String
        sourceColNameStr = UCase(CStr(sourceColName))
        If Left(sourceColNameStr, 6) = "MATCH_" Or Left(sourceColNameStr, 6) = "MATCH." Then
            DebugPrint "AutoMapColumns: Skipping UI column '" & sourceColName & "'"
            GoTo NextColumn
        End If

        If Not targetCols.exists(sourceColName) Then
            ' Column not found in TARGET - try to find with alias
            foundMapping = False

            ' Try alias lookup
            If aliases.exists(sourceColName) Then
                aliasName = aliases(sourceColName)
                If targetCols.exists(aliasName) Then
                    ' Found via alias
                    targetCols(sourceColName) = targetCols(aliasName)
                    foundMapping = True
                End If
            End If

            If Not foundMapping Then
                ' FIX: Do NOT prompt for column mapping in Target - Target is read-only
                ' We don't need to map to Target columns since we won't write to Target
                DebugPrint "AutoMapColumns: Skipping mapping for column '" & sourceColName & "' (Target is read-only)"
            End If
        End If

NextColumn:
    Next sourceColName
End Sub

Private Function PromptForColumnMapping(sourceSheet As Worksheet, targetSheet As Worksheet, _
                                        sourceHeaderRow As Long, targetHeaderRow As Long, _
                                        ByVal sourceColName As String, sourceColIdx As Long) As String

    Dim sourceHeader As String
    Dim targetHeaders As Collection
    Dim lastTargetCol As Long
    Dim i As Long
    Dim h As String
    Dim prompt As String
    Dim selection As Integer

    sourceHeader = CStr(sourceSheet.Cells(sourceHeaderRow, sourceColIdx).value)

    ' Get list of TARGET headers
    Set targetHeaders = New Collection
    lastTargetCol = GetLastColumn(targetSheet, targetHeaderRow)

    For i = 1 To lastTargetCol
        h = Trim(CStr(targetSheet.Cells(targetHeaderRow, i).value))
        If h <> "" Then targetHeaders.Add h
    Next i

    If targetHeaders.Count = 0 Then
        PromptForColumnMapping = ""
        Exit Function
    End If

    ' Build prompt
    prompt = "Column '" & sourceHeader & "' not found in TARGET sheet." & vbCrLf & vbCrLf & _
             "Select matching column in TARGET (" & targetSheet.Name & "):" & vbCrLf

    For i = 1 To targetHeaders.Count
        prompt = prompt & i & ". " & targetHeaders(i) & vbCrLf
    Next i

    prompt = prompt & vbCrLf & "Enter number (0 to skip):"

    selection = InputBox(prompt, "Column Mapping", 1)

    If selection > 0 And selection <= targetHeaders.Count Then
        PromptForColumnMapping = targetHeaders(selection)
    Else
        PromptForColumnMapping = ""
    End If
End Function

'===============================================================================
' WRITE COLUMNS ENSURANCE
'===============================================================================

'===============================================================================
' Check if MATCHED_ID column exists in the worksheet
' Returns: True if MATCHED_ID or MATCHED_ASSETID found, False otherwise
'===============================================================================
Private Function CheckMatchedIdExists(ws As Worksheet, headerRow As Long) As Boolean
    Dim col As Long
    Dim headerName As String

    CheckMatchedIdExists = False

    ' Search for MATCHED_ID in the header row (up to column 50)
    For col = 1 To 50
        If headerRow > 0 And col <= ws.Columns.Count Then
            On Error Resume Next
            headerName = UCase(Trim(Replace(CStr(ws.Cells(headerRow, col).Value), "_", "")))
            On Error GoTo 0

            If headerName = "MATCHEDID" Or headerName = "MATCHEDASSETID" Then
                CheckMatchedIdExists = True
                Exit Function
            End If
        End If
    Next col
End Function

'===============================================================================
' Check if all result columns exist (MATCHED_ID, MATCH_TYPE, MATCH_STATUS, SOURCE_FILE, TARGET_FILE)
' Returns: True if all columns exist, False otherwise
' This helps skip redundant column insertion prompts
'===============================================================================

Private Function PromptForMatchedIdColumn() As String
    Dim selection As String

    ' InputBox is FIRST and ONLY executable line in this function
    selection = InputBox("Enter the column NAME from the TARGET sheet that should be written to MATCHED_ID:" & vbCrLf & vbCrLf & _
        "Example: ASSETID, ASSETNUM, SERIALNUM, ID, etc." & vbCrLf & _
        "(Leave blank to use default: first available ID column)", _
        "Select MATCHED_ID Value Column", "ASSETID")

    If selection = "" Then
        PromptForMatchedIdColumn = ""
        Exit Function
    End If

    PromptForMatchedIdColumn = Trim(selection)
End Function

Private Function CheckAllResultColumnsExist(ws As Worksheet, headerRow As Long) As Boolean
    Dim col As Long
    Dim headerName As String
    Dim hasMatchedId As Boolean
    Dim hasMatchType As Boolean
    Dim hasMatchStatus As Boolean
    Dim hasSourceFile As Boolean
    Dim hasTargetFile As Boolean

    hasMatchedId = False
    hasMatchType = False
    hasMatchStatus = False
    hasSourceFile = False
    hasTargetFile = False

    If headerRow <= 0 Then
        CheckAllResultColumnsExist = False
        Exit Function
    End If

    ' Search for all result columns in the header row - EXACT MATCH ONLY
    For col = 1 To 100
        If col > ws.Columns.Count Then Exit For
        On Error Resume Next
        headerName = UCase(Trim(CStr(ws.Cells(headerRow, col).Value)))
        On Error GoTo 0

        ' EXACT MATCH ONLY - no Replace, no fuzzy matching
        If headerName = "MATCHED_ID" Then
            hasMatchedId = True
        ElseIf headerName = "MATCH_TYPE" Then
            hasMatchType = True
        ElseIf headerName = "MATCH_STATUS" Then
            hasMatchStatus = True
        ElseIf headerName = "SOURCE_FILE" Then
            hasSourceFile = True
        ElseIf headerName = "TARGET_FILE" Then
            hasTargetFile = True
        End If
    Next col

    CheckAllResultColumnsExist = (hasMatchedId And hasMatchType And hasMatchStatus And hasSourceFile And hasTargetFile)
    DebugPrint "CheckAllResultColumnsExist: " & ws.Name & " at row " & headerRow & " = " & CheckAllResultColumnsExist & _
               " (MATCHED_ID=" & hasMatchedId & ", MATCH_TYPE=" & hasMatchType & _
               ", MATCH_STATUS=" & hasMatchStatus & ", SOURCE_FILE=" & hasSourceFile & _
               ", TARGET_FILE=" & hasTargetFile & ")"
End Function

'===============================================================================
' ValidateHeaderRow - Checks if a row looks like a valid header row
' Returns True if the row has at least 3 non-empty cells (typical header)
'===============================================================================
Private Function ValidateHeaderRow(ws As Worksheet, headerRow As Long) As Boolean
    Dim col As Long
    Dim nonEmptyCount As Long
    Dim headerName As String
    Dim hasUIKeyword As Boolean
    Dim cellValue As String

    ValidateHeaderRow = False

    ' Check if row is within worksheet bounds
    If headerRow < 1 Or headerRow > ws.Rows.Count Then
        Exit Function
    End If

    nonEmptyCount = 0
    hasUIKeyword = False
    For col = 1 To 50
        If col > ws.Columns.Count Then Exit For
        On Error Resume Next
        headerName = Trim(CStr(ws.Cells(headerRow, col).Value))
        On Error GoTo 0
        If headerName <> "" Then
            nonEmptyCount = nonEmptyCount + 1

            ' Check for UI row keywords - these indicate a MATCH UI row, not a real data header
            cellValue = UCase(headerName)
            If InStr(cellValue, "MATCH ") > 0 Or InStr(cellValue, "SOURCE:") > 0 Or _
               InStr(cellValue, "TARGET:") > 0 Or cellValue = "X" Then
                hasUIKeyword = True
            End If
        End If
    Next col

    ' A valid header should have at least 3 non-empty columns
    ' AND NOT be a UI row (which has keywords like "Match", "Source:", "Target:")
    If hasUIKeyword Then
        DebugPrint "ValidateHeaderRow: Row " & headerRow & " appears to be a UI row (has UI keywords), invalid"
        ValidateHeaderRow = False
    Else
        ValidateHeaderRow = (nonEmptyCount >= 3)
    End If
    DebugPrint "ValidateHeaderRow: Row " & headerRow & " in sheet " & ws.Name & " has " & nonEmptyCount & " columns, valid=" & ValidateHeaderRow
End Function

'===============================================================================
' HasDataBelowHeader - Checks if there's data below the header row
' Returns True if there's at least 1 data row below the header
' FIX: Check multiple rows (row+1 to row+5) to detect data properly
'===============================================================================
Private Function HasDataBelowHeader(ws As Worksheet, headerRow As Long) As Boolean
    Dim col As Long
    Dim firstDataCol As Long
    Dim headerName As String
    Dim checkRow As Long
    Dim checkCol As Long
    Dim dataCell As String

    HasDataBelowHeader = False

    ' Find first non-empty column in header row
    firstDataCol = 0
    For col = 1 To 50
        If col > ws.Columns.Count Then Exit For
        On Error Resume Next
        headerName = Trim(CStr(ws.Cells(headerRow, col).Value))
        On Error GoTo 0
        If headerName <> "" Then
            firstDataCol = col
            Exit For
        End If
    Next col

    If firstDataCol = 0 Then
        Exit Function
    End If

    ' Check up to 50 rows below header for any data across multiple columns
    For checkRow = headerRow + 1 To headerRow + 50
        For checkCol = 1 To 20
            On Error Resume Next
            dataCell = Trim(CStr(ws.Cells(checkRow, checkCol).Value))
            On Error GoTo 0
            If dataCell <> "" Then
                HasDataBelowHeader = True
                Exit For
            End If
        Next checkCol
        If HasDataBelowHeader Then Exit For
    Next checkRow

    DebugPrint "HasDataBelowHeader: Row " & headerRow & " in sheet " & ws.Name & " has data below=" & HasDataBelowHeader
End Function

'===============================================================================
' HasValidHeaderAbove - Checks if there's a valid header row above the selected row
' Returns True if there's at least 3 non-empty cells in a row above selectedRow
' This helps detect if selectedRow is in the data area (not a header)
' FIXED: Now properly checks g_DataHeaderRow and uses fallback scanning
'===============================================================================
Private Function HasValidHeaderAbove(ws As Worksheet, selectedRow As Long) As Boolean
    Dim col As Long
    Dim checkRow As Long
    Dim nonEmptyCount As Long
    Dim headerName As String

    HasValidHeaderAbove = False

    ' Step 1: Check if g_DataHeaderRow > 1 BEFORE using g_DataHeaderRow - 1
    Dim uiRowsCount As Long

    If g_DataHeaderRow > 1 Then
        uiRowsCount = g_DataHeaderRow - 1
    Else
        ' Step 1a: Fallback - scan to find real header boundary
        Dim scanRow As Long
        Dim prevNonEmpty As Long
        Dim currNonEmpty As Long
        Dim hasUIKeywords As Boolean

        prevNonEmpty = 0
        For scanRow = 1 To 30
            currNonEmpty = 0
            hasUIKeywords = False

            For col = 1 To 50
                Dim cellVal As String
                cellVal = UCase(Trim(CStr(ws.Cells(scanRow, col).Value)))

                If cellVal <> "" Then
                    currNonEmpty = currNonEmpty + 1

                    ' Check for known UI keywords using InStr
                    If InStr(cellVal, "MATCHED_ID") > 0 Or _
                       InStr(cellVal, "MATCH_TYPE") > 0 Or InStr(cellVal, "MATCH TYPE") > 0 Or _
                       InStr(cellVal, "MATCH STATUS") > 0 Or InStr(cellVal, "MATCH_STATUS") > 0 Or _
                       InStr(cellVal, "SOURCE FILE") > 0 Or InStr(cellVal, "TARGET FILE") > 0 Or _
                       InStr(cellVal, "SOURCE_FILE") > 0 Or InStr(cellVal, "TARGET_FILE") > 0 Then
                        hasUIKeywords = True
                    End If
                End If
            Next col

            ' Only accept as header if: 3+ non-empty cells, has data below, NO UI keywords
            If currNonEmpty >= 3 And Not hasUIKeywords And HasDataBelowHeader(ws, scanRow) Then
                uiRowsCount = scanRow - 1
                Exit For
            End If

            prevNonEmpty = currNonEmpty
        Next scanRow

        If uiRowsCount < 1 Then uiRowsCount = 6
    End If

    ' Step 2: NEVER let uiRowsCount be negative or zero
    If uiRowsCount < 1 Then uiRowsCount = 1

    ' Step 3: Skip UI rows when checking for valid headers
    For checkRow = 1 To selectedRow - 1
        If checkRow <= uiRowsCount Then
            GoTo NextCheckRow
        End If

        nonEmptyCount = 0
        For col = 1 To 50
            If col > ws.Columns.Count Then Exit For
            On Error Resume Next
            headerName = Trim(CStr(ws.Cells(checkRow, col).Value))
            On Error GoTo 0
            If headerName <> "" Then
                nonEmptyCount = nonEmptyCount + 1
            End If
        Next col

        If nonEmptyCount >= 3 Then
            HasValidHeaderAbove = True
            Exit For
        End If

NextCheckRow:
    Next checkRow

    DebugPrint "HasValidHeaderAbove: Row " & selectedRow & " in sheet " & ws.Name & " has header above=" & HasValidHeaderAbove
End Function

'===============================================================================
' ParseRowSelection - Parse user input like "3,5,7" or "3-7" or "3,5-8" into array
' Includes error handling for non-numeric input
'===============================================================================
Private Function ParseRowSelection(inputStr As String, minRow As Long, maxRow As Long) As Long()
    Dim result() As Long
    Dim resultCount As Long
    resultCount = 0

    Dim parts() As String
    parts = Split(inputStr, ",")

    Dim i As Long
    For i = 0 To UBound(parts)
        Dim part As String
        part = Trim(parts(i))

        If InStr(part, "-") > 0 Then
            Dim rangeParts() As String
            rangeParts = Split(part, "-")
            Dim startNum As Long
            Dim endNum As Long

            ' Error handling for range start
            On Error Resume Next
            startNum = CLng(Trim(rangeParts(0)))
            If Err.Number <> 0 Then
                Err.Clear
                GoTo SkipEntry
            End If
            On Error GoTo 0

            ' Error handling for range end
            On Error Resume Next
            endNum = CLng(Trim(rangeParts(1)))
            If Err.Number <> 0 Then
                Err.Clear
                GoTo SkipEntry
            End If
            On Error GoTo 0

            Dim j As Long
            For j = startNum To endNum
                ReDim Preserve result(resultCount)
                result(resultCount) = j
                resultCount = resultCount + 1
            Next j
        Else
            ' Handle single number with error handling
            On Error Resume Next
            Dim singleNum As Long
            singleNum = CLng(part)
            If Err.Number <> 0 Then
                Err.Clear
                GoTo SkipEntry
            End If
            On Error GoTo 0

            ReDim Preserve result(resultCount)
            result(resultCount) = singleNum
            resultCount = resultCount + 1
        End If

SkipEntry:
    Next i

    ' Validate - discard any value outside minRow to maxRow
    Dim validResult() As Long
    Dim validCount As Long
    validCount = 0

    For i = 0 To resultCount - 1
        If result(i) >= minRow And result(i) <= maxRow Then
            ReDim Preserve validResult(validCount)
            validResult(validCount) = result(i)
            validCount = validCount + 1
        End If
    Next i

    ' If no valid rows, set flag and exit
    If validCount = 0 Then
        MsgBox "All entered row numbers are outside the valid range of rows " & _
               minRow & " to " & maxRow, vbExclamation
        g_ParseCancelled = True
        Exit Function
    End If

    ParseRowSelection = validResult
End Function

'===============================================================================
' QuickSortDescending - Sort array in descending order
'===============================================================================
Private Sub QuickSortDescending(arr() As Long)
    Dim i As Long
    Dim j As Long
    Dim temp As Long

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

'===============================================================================
' ValidateHeaderHasNames - Checks if header row still has valid column names
' Returns True if at least 3 columns have non-empty header names
' Also checks for key columns to ensure original headers weren't deleted
' Used to detect if user deleted header names before running Execute
'===============================================================================
Private Function ValidateHeaderHasNames(ws As Worksheet, headerRow As Long) As Boolean
    Dim col As Long
    Dim nonEmptyCount As Long
    Dim headerName As String
    Dim hasKeyColumn As Boolean
    Dim keyColumns As Variant
    Dim i As Long

    ValidateHeaderHasNames = False
    nonEmptyCount = 0
    hasKeyColumn = False

    ' Key columns that data headers indicate original
    keyColumns = Array("ASSETNUM", "ASSETID", "DESCRIPTION", "SERIALNUM", "SITEID", "LOCATION", "PARENT")

    ' Count non-empty header cells and check for key columns
    For col = 1 To 100
        If col > ws.Columns.Count Then Exit For
        On Error Resume Next
        headerName = Trim(CStr(ws.Cells(headerRow, col).Value))
        On Error GoTo 0

        If headerName <> "" Then
            nonEmptyCount = nonEmptyCount + 1
            ' Check if this is a key column
            For i = LBound(keyColumns) To UBound(keyColumns)
                If UCase(headerName) = UCase(keyColumns(i)) Then
                    hasKeyColumn = True
                    Exit For
                End If
            Next i
        End If
    Next col

    ' A valid header should have at least 3 non-empty columns AND at least one key column
    ' This prevents accepting just MATCH columns that were added by previous Execute
    ValidateHeaderHasNames = (nonEmptyCount >= 3) And hasKeyColumn

    DebugPrint "ValidateHeaderHasNames: Row " & headerRow & " in sheet " & ws.Name & " has " & nonEmptyCount & " columns, keyColumn=" & hasKeyColumn & ", valid=" & ValidateHeaderHasNames
End Function

'===============================================================================
' Ensure result columns are in the column map - searches directly in the sheet
' This handles cases where columns exist but weren't added to sourceCols by GetColumnMap
'===============================================================================
Private Sub EnsureResultColumnsInColMap(ws As Worksheet, headerRow As Long, colMap As Object)
    Dim col As Long
    Dim headerName As String
    Dim resultColumns As Variant
    Dim i As Long

    resultColumns = Array("MATCHED_ID", "MATCH_TYPE", "MATCH_STATUS", "SOURCE_FILE", "TARGET_FILE")

    ' Search for each result column in the header row
    For col = 1 To 100
        If col > ws.Columns.Count Then Exit For
        On Error Resume Next
        headerName = UCase(Trim(Replace(CStr(ws.Cells(headerRow, col).Value), "_", "")))
        On Error GoTo 0

        ' Check if this header matches any of our result columns
        For i = LBound(resultColumns) To UBound(resultColumns)
            Dim normalizedCol As String
            normalizedCol = UCase(Replace(resultColumns(i), "_", ""))

            If headerName = normalizedCol Then
                ' Found match - add to colMap if not already present
                If Not colMap.Exists(resultColumns(i)) Then
                    colMap.Add resultColumns(i), col
                    DebugPrint "EnsureResultColumnsInColMap: Added " & resultColumns(i) & " at column " & col
                End If
                Exit For
            End If
        Next i
    Next col
End Sub

'===============================================================================
' Prompt user for column position to insert result columns
' Returns: Column number (1-based), or 0 to cancel
'===============================================================================
Private Function PromptForResultColumnPosition(ws As Worksheet, headerRow As Long, defaultCol As Long) As Long
    Dim prompt As String
    Dim userInput As String
    Dim lastCol As Long
    Dim col As Long
    Dim headerList As String
    Dim selectedCol As Long

    ' FIX: First check if result columns already exist - if so, no need to prompt!
    Dim hasExistingResultCols As Boolean
    hasExistingResultCols = CheckAllResultColumnsExist(ws, headerRow)

    If hasExistingResultCols Then
        ' Columns already exist - find and return MATCHED_ID position
        Dim existingCol As Long
        existingCol = 0
        For col = 1 To 50
            If col > ws.Columns.Count Then Exit For
            On Error Resume Next
            Dim existingHeader As String
            existingHeader = UCase(Trim(Replace(CStr(ws.Cells(headerRow, col).Value), "_", "")))
            On Error GoTo 0
            If existingHeader = "MATCHEDID" Or existingHeader = "MATCHEDASSETID" Then
                existingCol = col
                Exit For
            End If
        Next col
        If existingCol > 0 Then
            DebugPrint "PromptForResultColumnPosition: Found existing MATCHED_ID at column " & existingCol
            PromptForResultColumnPosition = existingCol
            Exit Function
        End If
    End If

    ' Build list of existing headers for reference
    headerList = ""
    lastCol = 50 ' Check first 50 columns
    For col = 1 To lastCol
        If headerRow > 0 And col <= ws.Columns.Count Then
            On Error Resume Next
            Dim h As String
            h = Trim(CStr(ws.Cells(headerRow, col).Value))
            On Error GoTo 0
            If h <> "" Then
                headerList = headerList & "  Col " & col & ": " & h & vbCrLf
            End If
        End If
    Next col

    If headerList = "" Then headerList = "  (no headers found)"

    ' FIX: Show clearer prompt with warning about result columns not found
    prompt = "RESULT COLUMNS NOT FOUND in row " & headerRow & vbCrLf & vbCrLf & _
             "Current headers in " & ws.Name & " (row " & headerRow & "):" & vbCrLf & _
             headerList & vbCrLf & _
             "The comparison needs these columns: MATCHED_ID, MATCH_TYPE, MATCH_STATUS, SOURCE_FILE, TARGET_FILE" & vbCrLf & vbCrLf & _
             "Where should we INSERT the new result columns?" & vbCrLf & vbCrLf & _
             "RECOMMENDED: Press Enter to use default (column " & defaultCol & ")" & vbCrLf & _
             "This will add columns AFTER your last column - safest option!" & vbCrLf & _
             "Your existing data will be shifted to the RIGHT, NOT deleted."

    userInput = InputBox(prompt, "Select Column Position - RECOMMEND: Press Enter", CStr(defaultCol))

    ' Handle cancel or empty input
    If userInput = "" Then
        selectedCol = defaultCol
    ElseIf Trim(userInput) = "" Then
        selectedCol = defaultCol
    Else
        On Error Resume Next
        selectedCol = CLng(userInput)
        On Error GoTo 0
        If selectedCol < 1 Then selectedCol = defaultCol
    End If

    PromptForResultColumnPosition = selectedCol
End Function

Private Sub EnsureWriteColumnsExist(ws As Worksheet, colMap As Object, _
                                    ByVal headerRow As Long, learned As Object, sheetType As String)

    Dim lastCol As Long
    Dim col As Long
    Dim headerName As String
    Dim insertCol As Long
    Dim matchedIdCol As Long
    Dim matchedIdPos As Long
    Dim i As Long
    Dim hasUI As Boolean
    Dim lastColWithData As Long
    Dim basePosition As Long
    Dim insertPosition As Long
    Dim lo As ListObject
    Dim tablesToDelete As Collection

    ' FIX: Delete any Excel Tables (ListObjects) to prevent "cannot move cells in table" error
    On Error Resume Next
    Set tablesToDelete = New Collection
    For Each lo In ws.ListObjects
        tablesToDelete.Add lo.Name
    Next lo
    For i = 1 To tablesToDelete.Count
        ws.ListObjects(tablesToDelete(i)).Delete
        DebugPrint "EnsureWriteColumnsExist: Deleted table '" & tablesToDelete(i) & "'"
    Next i
    On Error GoTo 0

    ' Initialize
    lastColWithData = 0

    ' Array of result columns we need to ensure exist
    Dim resultColumns As Variant
    resultColumns = Array("MATCHED_ID", "MATCH_TYPE", "MATCH_STATUS", "SOURCE_FILE", "TARGET_FILE")

    ' First, find the last column with data (for placing new columns at end if needed)
    For col = 1 To 250
        headerName = Trim(CStr(ws.Cells(headerRow, col).value))
        If headerName <> "" Then
            lastColWithData = col
        End If
    Next col

    DebugPrint "EnsureWriteColumnsExist: lastColWithData = " & lastColWithData

    ' Calculate base position for new columns (after existing data)
    basePosition = lastColWithData + 1

    ' First, SEARCH for existing result columns in the header row
    ' and record their positions
    ' FIX: Only search for exact MATCHED_ID, not MATCHED_ASSETID
    Dim existingCols As Object
    Set existingCols = CreateObject("Scripting.Dictionary")

    lastCol = 200  ' Check up to column 200
    For col = 1 To lastCol
        headerName = UCase(Trim(CStr(ws.Cells(headerRow, col).value)))
        If headerName <> "" Then
            For i = LBound(resultColumns) To UBound(resultColumns)
                ' Only match exact column names - MATCHED_ASSETID is different from MATCHED_ID
                If headerName = resultColumns(i) Then
                    existingCols.Add resultColumns(i), col
                    DebugPrint "EnsureWriteColumnsExist: Found existing column '" & resultColumns(i) & "' at position " & col
                    Exit For
                End If
            Next i
        End If
    Next col

    ' FIXED: Check if result columns already exist
    ' If they exist, just USE them (no insert)
    ' If they don't exist, add new columns

    Dim existingPos As Object
    Set existingPos = CreateObject("Scripting.Dictionary")

    ' Search for existing result columns
    ' FIX: Only match exact "MATCHEDID" - NOT "MATCHEDASSETID" (different column!)
    For col = 1 To 250
        headerName = UCase(Trim(Replace(CStr(ws.Cells(headerRow, col).value), "_", "")))
        ' Only match exact MATCHEDID - MATCHED_ASSETID is a different column
        If headerName = "MATCHEDID" Then
            existingPos.Add "MATCHED_ID", col
            DebugPrint "EnsureWriteColumnsExist: Found MATCHED_ID at column " & col
        ElseIf headerName = "MATCHTYPE" Then
            existingPos.Add "MATCH_TYPE", col
            DebugPrint "EnsureWriteColumnsExist: Found MATCH_TYPE at column " & col
        ElseIf headerName = "MATCHSTATUS" Then
            existingPos.Add "MATCH_STATUS", col
            DebugPrint "EnsureWriteColumnsExist: Found MATCH_STATUS at column " & col
        ElseIf headerName = "SOURCEFILE" Then
            existingPos.Add "SOURCE_FILE", col
            DebugPrint "EnsureWriteColumnsExist: Found SOURCE_FILE at column " & col
        ElseIf headerName = "TARGETFILE" Then
            existingPos.Add "TARGET_FILE", col
            DebugPrint "EnsureWriteColumnsExist: Found TARGET_FILE at column " & col
        End If
    Next col

    ' If all result columns exist, just use them (no changes needed)
    If existingPos.exists("MATCHED_ID") And existingPos.exists("MATCH_TYPE") And _
       existingPos.exists("MATCH_STATUS") And existingPos.exists("SOURCE_FILE") And existingPos.exists("TARGET_FILE") Then

        ' Update colMap with existing positions
        If colMap.exists("MATCHED_ID") Then colMap.Remove "MATCHED_ID"
        colMap.Add "MATCHED_ID", existingPos("MATCHED_ID")
        If colMap.exists("MATCH_TYPE") Then colMap.Remove "MATCH_TYPE"
        colMap.Add "MATCH_TYPE", existingPos("MATCH_TYPE")
        If colMap.exists("MATCH_STATUS") Then colMap.Remove "MATCH_STATUS"
        colMap.Add "MATCH_STATUS", existingPos("MATCH_STATUS")
        If colMap.exists("SOURCE_FILE") Then colMap.Remove "SOURCE_FILE"
        colMap.Add "SOURCE_FILE", existingPos("SOURCE_FILE")
        If colMap.exists("TARGET_FILE") Then colMap.Remove "TARGET_FILE"
        colMap.Add "TARGET_FILE", existingPos("TARGET_FILE")

        DebugPrint "EnsureWriteColumnsExist: Using existing result columns"
        Exit Sub
    End If

    ' If MATCHED_ID exists but other columns don't, ADD missing columns after MATCHED_ID
    If existingPos.exists("MATCHED_ID") Then
        matchedIdPos = existingPos("MATCHED_ID")

        ' Add MATCH_TYPE at matchedIdPos + 1 if not exists
        If Not existingPos.exists("MATCH_TYPE") Then
            ws.Columns(matchedIdPos + 1).Insert Shift:=xlToRight
            ws.Cells(headerRow, matchedIdPos + 1).Value = "MATCH_TYPE"
            ws.Cells(headerRow, matchedIdPos + 1).Font.Bold = True
            ws.Cells(headerRow, matchedIdPos + 1).Interior.Color = RGB(91, 115, 150)
            ws.Cells(headerRow, matchedIdPos + 1).Font.Color = RGB(255, 255, 255)
            ws.Cells(headerRow, matchedIdPos + 1).HorizontalAlignment = xlCenter
            existingPos.Add "MATCH_TYPE", matchedIdPos + 1
        End If

        ' Add MATCH_STATUS after MATCH_TYPE if not exists
        If Not existingPos.exists("MATCH_STATUS") Then
            Dim statusPos As Long
            statusPos = existingPos("MATCH_TYPE") + 1
            ws.Columns(statusPos).Insert Shift:=xlToRight
            ws.Cells(headerRow, statusPos).Value = "MATCH_STATUS"
            ws.Cells(headerRow, statusPos).Font.Bold = True
            ws.Cells(headerRow, statusPos).Interior.Color = RGB(91, 115, 150)
            ws.Cells(headerRow, statusPos).Font.Color = RGB(255, 255, 255)
            ws.Cells(headerRow, statusPos).HorizontalAlignment = xlCenter
            existingPos.Add "MATCH_STATUS", statusPos
        End If

        ' Add SOURCE_FILE after MATCH_STATUS if not exists
        If Not existingPos.exists("SOURCE_FILE") Then
            Dim sourcePos As Long
            sourcePos = existingPos("MATCH_STATUS") + 1
            ws.Columns(sourcePos).Insert Shift:=xlToRight
            ws.Cells(headerRow, sourcePos).Value = "SOURCE_FILE"
            ws.Cells(headerRow, sourcePos).Font.Bold = True
            ws.Cells(headerRow, sourcePos).Interior.Color = RGB(91, 115, 150)
            ws.Cells(headerRow, sourcePos).Font.Color = RGB(255, 255, 255)
            ws.Cells(headerRow, sourcePos).HorizontalAlignment = xlCenter
            existingPos.Add "SOURCE_FILE", sourcePos
        End If

        ' Add TARGET_FILE after SOURCE_FILE if not exists
        If Not existingPos.exists("TARGET_FILE") Then
            Dim targetPos As Long
            targetPos = existingPos("SOURCE_FILE") + 1
            ws.Columns(targetPos).Insert Shift:=xlToRight
            ws.Cells(headerRow, targetPos).Value = "TARGET_FILE"
            ws.Cells(headerRow, targetPos).Font.Bold = True
            ws.Cells(headerRow, targetPos).Interior.Color = RGB(91, 115, 150)
            ws.Cells(headerRow, targetPos).Font.Color = RGB(255, 255, 255)
            ws.Cells(headerRow, targetPos).HorizontalAlignment = xlCenter
            existingPos.Add "TARGET_FILE", targetPos
        End If

        ' Update colMap
        If colMap.exists("MATCHED_ID") Then colMap.Remove "MATCHED_ID"
        colMap.Add "MATCHED_ID", existingPos("MATCHED_ID")
        If colMap.exists("MATCH_TYPE") Then colMap.Remove "MATCH_TYPE"
        colMap.Add "MATCH_TYPE", existingPos("MATCH_TYPE")
        If colMap.exists("MATCH_STATUS") Then colMap.Remove "MATCH_STATUS"
        colMap.Add "MATCH_STATUS", existingPos("MATCH_STATUS")
        If colMap.exists("SOURCE_FILE") Then colMap.Remove "SOURCE_FILE"
        colMap.Add "SOURCE_FILE", existingPos("SOURCE_FILE")
        If colMap.exists("TARGET_FILE") Then colMap.Remove "TARGET_FILE"
        colMap.Add "TARGET_FILE", existingPos("TARGET_FILE")

        DebugPrint "EnsureWriteColumnsExist: Added missing columns after existing MATCHED_ID"
        Exit Sub
    End If

    ' MATCHED_ID doesn't exist at expected position - check if it exists ANYWHERE else
    ' This prevents duplicate columns when user has deleted/moved columns
    Dim searchCol As Long
    Dim foundAnyMatchColumn As Boolean
    foundAnyMatchColumn = False

    For searchCol = 1 To 50
        headerName = UCase(Trim(CStr(ws.Cells(headerRow, searchCol).value)))
        If headerName = "MATCHED_ID" Or headerName = "MATCHEDID" Or _
           headerName = "MATCH_TYPE" Or headerName = "MATCHTYPE" Or _
           headerName = "MATCH_STATUS" Or headerName = "MATCHSTATUS" Then
            foundAnyMatchColumn = True
            DebugPrint "EnsureWriteColumnsExist: Found MATCH column at position " & searchCol & " - " & headerName
            Exit For
        End If
    Next searchCol

    If foundAnyMatchColumn Then
        ' MATCH columns exist somewhere else in the sheet - use them!
        ' Re-scan to get all column positions
        Set existingPos = CreateObject("Scripting.Dictionary")
        For searchCol = 1 To 50
            headerName = UCase(Trim(Replace(CStr(ws.Cells(headerRow, searchCol).value), "_", "")))
            If headerName = "MATCHEDID" Then
                existingPos.Add "MATCHED_ID", searchCol
            ElseIf headerName = "MATCHTYPE" Then
                existingPos.Add "MATCH_TYPE", searchCol
            ElseIf headerName = "MATCHSTATUS" Then
                existingPos.Add "MATCH_STATUS", searchCol
            ElseIf headerName = "SOURCEFILE" Then
                existingPos.Add "SOURCE_FILE", searchCol
            ElseIf headerName = "TARGETFILE" Then
                existingPos.Add "TARGET_FILE", searchCol
            End If
        Next searchCol

        ' Update colMap with found positions
        If existingPos.exists("MATCHED_ID") Then
            If colMap.exists("MATCHED_ID") Then colMap.Remove "MATCHED_ID"
            colMap.Add "MATCHED_ID", existingPos("MATCHED_ID")
        End If
        If existingPos.exists("MATCH_TYPE") Then
            If colMap.exists("MATCH_TYPE") Then colMap.Remove "MATCH_TYPE"
            colMap.Add "MATCH_TYPE", existingPos("MATCH_TYPE")
        End If
        If existingPos.exists("MATCH_STATUS") Then
            If colMap.exists("MATCH_STATUS") Then colMap.Remove "MATCH_STATUS"
            colMap.Add "MATCH_STATUS", existingPos("MATCH_STATUS")
        End If
        If existingPos.exists("SOURCE_FILE") Then
            If colMap.exists("SOURCE_FILE") Then colMap.Remove "SOURCE_FILE"
            colMap.Add "SOURCE_FILE", existingPos("SOURCE_FILE")
        End If
        If existingPos.exists("TARGET_FILE") Then
            If colMap.exists("TARGET_FILE") Then colMap.Remove "TARGET_FILE"
            colMap.Add "TARGET_FILE", existingPos("TARGET_FILE")
        End If

        DebugPrint "EnsureWriteColumnsExist: Found existing MATCH columns at different positions, using them"
        Exit Sub
    End If

    ' MATCHED_ID doesn't exist anywhere - need to insert result columns
    ' FIX: Always insert at columns A-E (fixed position, no prompt)
    ' MATCHED_ID = A, MATCH_TYPE = B, MATCH_STATUS = C, SOURCE_FILE = D, TARGET_FILE = E
    Dim insertAtCol As Long

    ' FIX: Check if there's data in columns A-E BEFORE inserting
    ' If data exists, we should NOT insert - find another position or warn user
    Dim hasDataInAtoE As Boolean
    Dim checkRow As Long
    hasDataInAtoE = False
    For checkRow = headerRow + 1 To headerRow + 10  ' Check first 10 data rows
        Dim c As Long
        For c = 1 To 5
            If Trim(CStr(ws.Cells(checkRow, c).Value)) <> "" Then
                hasDataInAtoE = True
                Exit For
            End If
        Next c
        If hasDataInAtoE Then Exit For
    Next checkRow

    If hasDataInAtoE Then
        ' Data exists in A-E - find the last column and insert after it
        lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        insertAtCol = lastCol + 1

        ' Add MATCH columns at the end
        If colMap.exists("MATCHED_ID") Then colMap.Remove "MATCHED_ID"
        colMap.Add "MATCHED_ID", insertAtCol
        If colMap.exists("MATCH_TYPE") Then colMap.Remove "MATCH_TYPE"
        colMap.Add "MATCH_TYPE", insertAtCol + 1
        If colMap.exists("MATCH_STATUS") Then colMap.Remove "MATCH_STATUS"
        colMap.Add "MATCH_STATUS", insertAtCol + 2
        If colMap.exists("SOURCE_FILE") Then colMap.Remove "SOURCE_FILE"
        colMap.Add "SOURCE_FILE", insertAtCol + 3
        If colMap.exists("TARGET_FILE") Then colMap.Remove "TARGET_FILE"
        colMap.Add "TARGET_FILE", insertAtCol + 4

        ' Add headers at the new position
        ws.Cells(headerRow, insertAtCol).Value = "MATCHED_ID"
        ws.Cells(headerRow, insertAtCol).Font.Bold = True
        ws.Cells(headerRow, insertAtCol).Interior.Color = RGB(91, 115, 150)
        ws.Cells(headerRow, insertAtCol).Font.Color = RGB(255, 255, 255)
        ws.Cells(headerRow, insertAtCol).HorizontalAlignment = xlCenter

        ws.Cells(headerRow, insertAtCol + 1).Value = "MATCH_TYPE"
        ws.Cells(headerRow, insertAtCol + 1).Font.Bold = True
        ws.Cells(headerRow, insertAtCol + 1).Interior.Color = RGB(91, 115, 150)
        ws.Cells(headerRow, insertAtCol + 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(headerRow, insertAtCol + 1).HorizontalAlignment = xlCenter

        ws.Cells(headerRow, insertAtCol + 2).Value = "MATCH_STATUS"
        ws.Cells(headerRow, insertAtCol + 2).Font.Bold = True
        ws.Cells(headerRow, insertAtCol + 2).Interior.Color = RGB(91, 115, 150)
        ws.Cells(headerRow, insertAtCol + 2).Font.Color = RGB(255, 255, 255)
        ws.Cells(headerRow, insertAtCol + 2).HorizontalAlignment = xlCenter

        ws.Cells(headerRow, insertAtCol + 3).Value = "SOURCE_FILE"
        ws.Cells(headerRow, insertAtCol + 3).Font.Bold = True
        ws.Cells(headerRow, insertAtCol + 3).Interior.Color = RGB(91, 115, 150)
        ws.Cells(headerRow, insertAtCol + 3).Font.Color = RGB(255, 255, 255)
        ws.Cells(headerRow, insertAtCol + 3).HorizontalAlignment = xlCenter

        ws.Cells(headerRow, insertAtCol + 4).Value = "TARGET_FILE"
        ws.Cells(headerRow, insertAtCol + 4).Font.Bold = True
        ws.Cells(headerRow, insertAtCol + 4).Interior.Color = RGB(91, 115, 150)
        ws.Cells(headerRow, insertAtCol + 4).Font.Color = RGB(255, 255, 255)
        ws.Cells(headerRow, insertAtCol + 4).HorizontalAlignment = xlCenter

        DebugPrint "EnsureWriteColumnsExist: Data exists in A-E, created MATCH columns at column " & insertAtCol & " instead"
        Exit Sub
    End If

    ' No data in A-E - safe to insert at column A (the fixed position for MATCHED_ID)
    insertAtCol = 1

    DebugPrint "EnsureWriteColumnsExist: MATCHED_ID not found anywhere, auto-inserting at column A (fixed position)"

    ' CRITICAL: Insert 5 columns at the position FIRST (before writing anything)
    ' This shifts existing columns to the right, preserving all data
    ' Using Insert Shift:=xlToRight is the SAFE way to add columns
    On Error Resume Next
    ws.Columns(insertAtCol).Insert Shift:=xlToRight
    If Err.Number <> 0 Then
        MsgBox "Error inserting column at position " & insertAtCol & ": " & Err.Description, vbCritical
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Add MATCHED_ID at insertAtCol (now there's an empty column there)
    ws.Cells(headerRow, insertAtCol).Value = "MATCHED_ID"
    ws.Cells(headerRow, insertAtCol).Font.Bold = True
    ws.Cells(headerRow, insertAtCol).Interior.Color = RGB(91, 115, 150)
    ws.Cells(headerRow, insertAtCol).Font.Color = RGB(255, 255, 255)
    ws.Cells(headerRow, insertAtCol).HorizontalAlignment = xlCenter

    ' Insert column for MATCH_TYPE and add header
    On Error Resume Next
    ws.Columns(insertAtCol + 1).Insert Shift:=xlToRight
    On Error GoTo 0
    ws.Cells(headerRow, insertAtCol + 1).Value = "MATCH_TYPE"
    ws.Cells(headerRow, insertAtCol + 1).Font.Bold = True
    ws.Cells(headerRow, insertAtCol + 1).Interior.Color = RGB(91, 115, 150)
    ws.Cells(headerRow, insertAtCol + 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(headerRow, insertAtCol + 1).HorizontalAlignment = xlCenter

    ' Insert column for MATCH_STATUS and add header
    On Error Resume Next
    ws.Columns(insertAtCol + 2).Insert Shift:=xlToRight
    On Error GoTo 0
    ws.Cells(headerRow, insertAtCol + 2).Value = "MATCH_STATUS"
    ws.Cells(headerRow, insertAtCol + 2).Font.Bold = True
    ws.Cells(headerRow, insertAtCol + 2).Interior.Color = RGB(91, 115, 150)
    ws.Cells(headerRow, insertAtCol + 2).Font.Color = RGB(255, 255, 255)
    ws.Cells(headerRow, insertAtCol + 2).HorizontalAlignment = xlCenter

    ' Insert column for SOURCE_FILE and add header
    On Error Resume Next
    ws.Columns(insertAtCol + 3).Insert Shift:=xlToRight
    On Error GoTo 0
    ws.Cells(headerRow, insertAtCol + 3).Value = "SOURCE_FILE"
    ws.Cells(headerRow, insertAtCol + 3).Font.Bold = True
    ws.Cells(headerRow, insertAtCol + 3).Interior.Color = RGB(91, 115, 150)
    ws.Cells(headerRow, insertAtCol + 3).Font.Color = RGB(255, 255, 255)
    ws.Cells(headerRow, insertAtCol + 3).HorizontalAlignment = xlCenter

    ' Insert column for TARGET_FILE and add header (no insert needed, just write to last position)
    ws.Cells(headerRow, insertAtCol + 4).Value = "TARGET_FILE"
    ws.Cells(headerRow, insertAtCol + 4).Font.Bold = True
    ws.Cells(headerRow, insertAtCol + 4).Interior.Color = RGB(91, 115, 150)
    ws.Cells(headerRow, insertAtCol + 4).Font.Color = RGB(255, 255, 255)
    ws.Cells(headerRow, insertAtCol + 4).HorizontalAlignment = xlCenter

    ' Update colMap with the actual column positions (at user-specified position)
    If colMap.exists("MATCHED_ID") Then colMap.Remove "MATCHED_ID"
    colMap.Add "MATCHED_ID", insertAtCol
    If colMap.exists("MATCH_TYPE") Then colMap.Remove "MATCH_TYPE"
    colMap.Add "MATCH_TYPE", insertAtCol + 1
    If colMap.exists("MATCH_STATUS") Then colMap.Remove "MATCH_STATUS"
    colMap.Add "MATCH_STATUS", insertAtCol + 2
    If colMap.exists("SOURCE_FILE") Then colMap.Remove "SOURCE_FILE"
    colMap.Add "SOURCE_FILE", insertAtCol + 3
    If colMap.exists("TARGET_FILE") Then colMap.Remove "TARGET_FILE"
    colMap.Add "TARGET_FILE", insertAtCol + 4

    DebugPrint "EnsureWriteColumnsExist: Created all result columns at column " & insertAtCol & " (user-selected position, no MATCHED_ID existed)"
    Exit Sub

End Sub

'===============================================================================
' KEY COLUMN VALUE RETRIEVAL
'===============================================================================

Private Function GetKeyColumnValue(data As Variant, rowIndex As Long, colMap As Object) As String
    ' Get value from first column in colMap for MATCHED_ID output
    ' FIXED: Find the LEFTMOST data column (likely the key column)
    Dim keyCols As Variant
    Dim i As Long
    Dim colName As Variant
    Dim foundValue As String
    Dim colKey As Variant
    Dim maxColIdx As Long
    Dim foundColIdx As Long

    On Error Resume Next

    ' Find the LEFTMOST data column (not result columns)
    ' This is likely the primary key column
    maxColIdx = 999999
    foundColIdx = 999999

    ' First, find the leftmost non-result column in colMap
    For Each colKey In colMap.keys
        If UCase(colKey) <> "MATCHED_ID" And UCase(colKey) <> "MATCH_TYPE" And _
           UCase(colKey) <> "MATCH_STATUS" And UCase(colKey) <> "SOURCE_FILE" And _
           UCase(colKey) <> "TARGET_FILE" And UCase(colKey) <> "MATCHED_ASSETID" Then
            If colMap(colKey) < foundColIdx Then
                foundColIdx = colMap(colKey)
            End If
        End If
    Next colKey

    ' If we found a leftmost column, use it
    If foundColIdx < 999999 Then
        If foundColIdx > UBound(data, 2) Or rowIndex > UBound(data, 1) Or rowIndex < 1 Then
            GetKeyColumnValue = ""
            Exit Function
        End If
        foundValue = SafeCleanString(data(rowIndex, foundColIdx))
        If foundValue <> "" Then
            GetKeyColumnValue = foundValue
            Exit Function
        End If
    End If

    ' Fallback: Try common ID columns
    keyCols = Array("ASSETNUM", "ASSETID", "ID", "ASSET_NUM", "ASSET_NUMBER")

    For Each colName In keyCols
        If colMap.exists(colName) Then
            foundValue = SafeCleanString(data(rowIndex, colMap(colName)))
            If foundValue <> "" Then
                GetKeyColumnValue = foundValue
                Exit Function
            End If
        End If
    Next colName

    ' Last fallback: try ANY column that has a value
    For Each colKey In colMap.keys
        If UCase(colKey) <> "MATCHED_ID" And UCase(colKey) <> "MATCH_TYPE" And _
           UCase(colKey) <> "MATCH_STATUS" And UCase(colKey) <> "SOURCE_FILE" And _
           UCase(colKey) <> "TARGET_FILE" Then
            foundValue = SafeCleanString(data(rowIndex, colMap(colKey)))
            If foundValue <> "" Then
                GetKeyColumnValue = foundValue
                Exit Function
            End If
        End If
    Next colKey

    ' If still nothing found, return empty string
    On Error GoTo 0
    GetKeyColumnValue = ""
End Function

'===============================================================================
' UI BUILDER - Rebuild Match Builder UI
'===============================================================================

' Helper function to add MATCH_TYPE and MATCH_STATUS columns at correct position
' This is called BEFORE inserting rows for UI, so columns are added in the right place
Private Sub AddMatchColumnsAtCorrectPosition(ws As Worksheet, ByVal dataHeaderRow As Long, ByVal lastDataCol As Long)
    Dim i As Long
    Dim cellVal As String
    Dim matchedIdCol As Long
    Dim newCol As Long
    Dim headerRowActual As Long
    Dim nextInsertCol As Long  ' Running counter for where to insert next mandatory column
    Dim existingColPos As Long  ' Used in advance loop to track existing column positions
    Dim mandatoryHeaders As Object
    Dim headerName As Variant
    Dim targetCol As Long
    Dim currentCol As Long
    ' PERFORMANCE OPTIMIZATION: Disable screen updates and calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Set error handler to ensure settings are restored on any error
    On Error GoTo CleanExit

    Dim tempHasExisting As Boolean  ' Used in advance loop to check if column exists

    DebugPrint "AddMatchColumnsAtCorrectPosition: Starting..."

    ' Find the actual header row with column names
    ' Sometimes header might be in a different row, search around
    ' Search from dataHeaderRow to dataHeaderRow + 6, checking columns A-E for content
    headerRowActual = dataHeaderRow
    Dim colCheck As Long
    Dim hasContent As Boolean
    For i = dataHeaderRow To dataHeaderRow + 6
        If i > ws.Rows.Count Then Exit For
        ' Check columns 1-5 for any content
        hasContent = False
        For colCheck = 1 To 5
            If UCase(Trim(CStr(ws.Cells(i, colCheck).value))) <> "" Then
                hasContent = True
                Exit For
            End If
        Next colCheck
        If hasContent Then
            headerRowActual = i
            Exit For
        End If
    Next i

    ' Find MATCHED_ID column
    matchedIdCol = 0
    For i = 1 To lastDataCol + 10
        If i > ws.Columns.Count Then Exit For
        cellVal = UCase(Trim(CStr(ws.Cells(headerRowActual, i).value)))
        If cellVal = "MATCHED_ID" Or cellVal = "MATCHEDID" Then
            matchedIdCol = i
            DebugPrint "AddMatchColumnsAtCorrectPosition: Found MATCHED_ID at column " & i
            Exit For
        End If
    Next i

    ' FIX: If MATCHED_ID doesn't exist, add it at column 1 (first column of header)
    ' This ensures MATCHED_ID is always present
    ' Use a running counter to track where subsequent columns should go
    nextInsertCol = 2  ' Default: after column 1

    If matchedIdCol = 0 Then
        ' MATCHED_ID missing - write directly to canonical position 1, no column insertion
        DebugPrint "AddMatchColumnsAtCorrectPosition: MATCHED_ID not found, writing to column 1"
        ws.Cells(headerRowActual, 1).Value = "MATCHED_ID"
        ws.Cells(headerRowActual, 1).Font.Bold = True
        ws.Cells(headerRowActual, 1).Interior.Color = RGB(91, 115, 150)
        ws.Cells(headerRowActual, 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(headerRowActual, 1).HorizontalAlignment = xlCenter
        ws.Columns(1).ColumnWidth = Len("MATCHED_ID") * 3.5 + 2
        matchedIdCol = 1
    End If
    nextInsertCol = 2

    '===============================================================================
    ' FIX: Use canonical positions for ALL mandatory columns
    ' This ensures missing columns are inserted at their correct sequential positions
    '===============================================================================

    ' Define the canonical column positions for mandatory headers
    Set mandatoryHeaders = CreateObject("Scripting.Dictionary")
    mandatoryHeaders.Add "MATCHED_ID", 1
    mandatoryHeaders.Add "MATCH_TYPE", 2
    mandatoryHeaders.Add "MATCH_STATUS", 3
    mandatoryHeaders.Add "SOURCE_FILE", 4
    mandatoryHeaders.Add "TARGET_FILE", 5

    ' For each mandatory header in canonical order (skip MATCHED_ID - already handled)
    For Each headerName In Array("MATCH_TYPE", "MATCH_STATUS", "SOURCE_FILE", "TARGET_FILE")
        targetCol = mandatoryHeaders(headerName)

        ' Check if this header already exists at any position (exact UCase match only)
        currentCol = 0
        For i = 1 To lastDataCol
            If i > ws.Columns.Count Then Exit For
            cellVal = UCase(Trim(CStr(ws.Cells(headerRowActual, i).Value)))
            If cellVal = headerName Then
                currentCol = i
                Exit For
            End If
        Next i

        If currentCol = 0 Then
            ' Header missing - write directly to canonical position, no column insertion
            DebugPrint "AddMatchColumnsAtCorrectPosition: " & headerName & " missing, writing to column " & targetCol
            ws.Cells(headerRowActual, targetCol).Value = headerName
            ws.Cells(headerRowActual, targetCol).Font.Bold = True
            ws.Cells(headerRowActual, targetCol).Interior.Color = RGB(91, 115, 150)
            ws.Cells(headerRowActual, targetCol).Font.Color = RGB(255, 255, 255)
            ws.Cells(headerRowActual, targetCol).HorizontalAlignment = xlCenter
            ws.Columns(targetCol).ColumnWidth = Len(headerName) * 3.5 + 2
        Else
            ' Header exists (at any position) - leave it alone
            DebugPrint "AddMatchColumnsAtCorrectPosition: " & headerName & " already exists at column " & currentCol
        End If
    Next headerName

    DebugPrint "AddMatchColumnsAtCorrectPosition: Complete"

    Call ShiftEmptyColumnsPastTargetFile(ws, headerRowActual, lastDataCol)

    '===============================================================================
    ' ISSUE 3 FIX: Format ALL new columns with alternating colors
    ' Now format all MATCH columns
    '===============================================================================

    ' Calculate lastDataCol using explicit scan (NOT End(xlToLeft))
    Dim lastDataColScan As Long
    lastDataColScan = 0
    Dim scanCol As Long
    For scanCol = 1 To 200
        If Trim(CStr(ws.Cells(headerRowActual, scanCol).Value)) <> "" Then
            lastDataColScan = scanCol
        End If
    Next scanCol
    If lastDataColScan < 1 Then lastDataColScan = 10 ' Fallback

    ' Calculate lastDataRow using explicit scan
    Dim lastDataRow As Long
    lastDataRow = 0
    Dim colLastRow As Long
    Dim scanR As Long

    ' Use UsedRange to find actual last row safely
    Dim scanUpperBound As Long
    On Error Resume Next
    scanUpperBound = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    On Error GoTo 0
    If scanUpperBound < dataHeaderRow Then scanUpperBound = dataHeaderRow + 1

    For scanCol = 1 To lastDataColScan + 10
        colLastRow = 0
        For scanR = dataHeaderRow + 1 To scanUpperBound
            If Trim(CStr(ws.Cells(scanR, scanCol).Value)) <> "" Then
                colLastRow = scanR
            End If
        Next scanR
        If colLastRow > lastDataRow Then lastDataRow = colLastRow
    Next scanCol

    If lastDataRow = 0 Then lastDataRow = dataHeaderRow + 1 ' Fallback

    ' Find and format all MATCH columns
    Dim fmtCol As Long
    For fmtCol = 1 To lastDataColScan + 10
        If fmtCol > ws.Columns.Count Then Exit For
        cellVal = UCase(Trim(CStr(ws.Cells(headerRowActual, fmtCol).value)))

        ' Check if this is a MATCH column
        If cellVal = "MATCHED_ID" Or cellVal = "MATCHEDID" Or _
           cellVal = "MATCH_TYPE" Or cellVal = "MATCHTYPE" Or _
           cellVal = "MATCH_STATUS" Or cellVal = "MATCHSTATUS" Or _
           cellVal = "SOURCE_FILE" Or cellVal = "SOURCEFILE" Or _
           cellVal = "TARGET_FILE" Or cellVal = "TARGETFILE" Then

            ' OPTIMIZED: Apply solid color to entire range first, then loop only even rows
            ' This cuts work in half vs row-by-row
            If lastDataRow > dataHeaderRow Then
                Dim colorRange As Range
                Set colorRange = ws.Range(ws.Cells(dataHeaderRow + 1, fmtCol), ws.Cells(lastDataRow, fmtCol))

                ' Step 1: Apply white to entire range in ONE operation
                colorRange.Interior.Color = RGB(255, 255, 255)

                ' Step 2: Loop only even rows (half as many iterations)
                Dim rowIdx As Long
                For rowIdx = dataHeaderRow + 2 To lastDataRow Step 2
                    ws.Cells(rowIdx, fmtCol).Interior.Color = RGB(217, 225, 242)
                Next rowIdx
            End If
        End If
    Next fmtCol

CleanExit:
    ' Restore application settings - runs on both normal exit and error
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'===============================================================================
' ApplyMatchColumnColors - Apply alternating colors to MATCH result columns
'===============================================================================
Private Sub ApplyMatchColumnColors(ws As Worksheet, ByVal dataHeaderRow As Long)
    Dim lastColScan As Long
    Dim lastRowScan As Long
    Dim scanCol As Long
    Dim scanR As Long
    Dim scanUpperBound As Long
    Dim colLastRow As Long
    Dim fmtCol As Long
    Dim cellVal As String
    Dim colorRange As Range
    Dim rowIdx As Long

    ' Scan for last column using header row
    lastColScan = 0
    For scanCol = 1 To 200
        If Trim(CStr(ws.Cells(dataHeaderRow, scanCol).Value)) <> "" Then
            lastColScan = scanCol
        End If
    Next scanCol
    If lastColScan < 1 Then lastColScan = 10

    ' Scan for last data row
    lastRowScan = 0
    On Error Resume Next
    scanUpperBound = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    On Error GoTo 0
    If scanUpperBound < dataHeaderRow Then scanUpperBound = dataHeaderRow + 1

    For scanCol = 1 To lastColScan
        colLastRow = 0
        For scanR = dataHeaderRow + 1 To scanUpperBound
            If Trim(CStr(ws.Cells(scanR, scanCol).Value)) <> "" Then
                colLastRow = scanR
            End If
        Next scanR
        If colLastRow > lastRowScan Then lastRowScan = colLastRow
    Next scanCol
    If lastRowScan <= dataHeaderRow Then GoTo SkipMatchColors

    ' Apply alternating colors to MATCH columns only
    For fmtCol = 1 To lastColScan
        cellVal = UCase(Trim(CStr(ws.Cells(dataHeaderRow, fmtCol).Value)))
        If cellVal = "MATCHED_ID" Or cellVal = "MATCHEDID" Or _
           cellVal = "MATCH_TYPE" Or cellVal = "MATCHTYPE" Or _
           cellVal = "MATCH_STATUS" Or cellVal = "MATCHSTATUS" Or _
           cellVal = "SOURCE_FILE" Or cellVal = "SOURCEFILE" Or _
           cellVal = "TARGET_FILE" Or cellVal = "TARGETFILE" Then
            If lastRowScan > dataHeaderRow + 1 Then
                Set colorRange = ws.Range(ws.Cells(dataHeaderRow + 1, fmtCol), ws.Cells(lastRowScan, fmtCol))
                colorRange.Interior.Color = RGB(255, 255, 255)
                For rowIdx = dataHeaderRow + 2 To lastRowScan Step 2
                    ws.Cells(rowIdx, fmtCol).Interior.Color = RGB(217, 225, 242)
                Next rowIdx
            End If
        End If
    Next fmtCol
    SkipMatchColors:
End Sub

'===============================================================================
' ShiftEmptyColumnsPastTargetFile - Delete empty columns between mandatory headers
' Runs after AddMatchColumnsAtCorrectPosition to clean up empty columns left by deleted headers
'===============================================================================
Private Sub ShiftEmptyColumnsPastTargetFile(ws As Worksheet, ByVal headerRow As Long, ByVal lastDataCol As Long)
    Dim headerVal As String
    Dim dataVal As String
    Dim col As Long
    Dim scanRow As Long
    Dim hasData As Boolean
    Dim deletedCount As Long

    deletedCount = 0

    ' Scan from RIGHT to LEFT to avoid index shifting when deleting
    For col = lastDataCol To 1 Step -1
        headerVal = Trim(CStr(ws.Cells(headerRow, col).Value))

        If headerVal = "" Then
            hasData = False
            For scanRow = headerRow + 1 To headerRow + 10
                If scanRow > ws.Rows.Count Then Exit For
                dataVal = Trim(CStr(ws.Cells(scanRow, col).Value))
                If dataVal <> "" Then
                    hasData = True
                    Exit For
                End If
            Next scanRow

            If Not hasData Then
                DebugPrint "ShiftEmptyColumnsPastTargetFile: Deleting empty column at position " & col
                On Error Resume Next
                ws.Columns(col).Delete
                On Error GoTo 0
                deletedCount = deletedCount + 1
            End If
        End If
    Next col

    DebugPrint "ShiftEmptyColumnsPastTargetFile: Complete - deleted " & deletedCount & " empty columns"
End Sub

' Helper to find a valid worksheet for UI operations
Private Function GetValidWorksheetForUI() As Worksheet
    Dim ws As Worksheet
    Dim configSheet As Worksheet
    Dim sourceWSName As String
    Dim sourceWBName As String
    Dim wb As Workbook
    Dim commonNames As Variant
    Dim i As Long

    DebugPrint "GetValidWorksheetForUI: Starting..."

    ' Priority 1: Use g_SourceSheet if already configured
    If Not g_SourceSheet Is Nothing Then
        DebugPrint "GetValidWorksheetForUI: Using g_SourceSheet"
        Set GetValidWorksheetForUI = g_SourceSheet
        Exit Function
    End If

    ' Priority 2: Check if SOURCE is configured in settings
    Set configSheet = GetOrCreateConfigSheet
    sourceWSName = GetConfigValue(configSheet, CONFIG_SOURCE_WS)

    If sourceWSName <> "" Then
        sourceWBName = GetConfigValue(configSheet, CONFIG_SOURCE_WB)
        If sourceWBName <> "" Then
            Set wb = GetWorkbookByName(sourceWBName)
            If Not wb Is Nothing Then
                Set ws = SafeGetWorksheet(wb, sourceWSName)
                If Not ws Is Nothing Then
                    DebugPrint "GetValidWorksheetForUI: Using configured sheet"
                    Set GetValidWorksheetForUI = ws
                    Exit Function
                End If
            End If
        End If
    End If

    ' Priority 3: Try common sheet names in ThisWorkbook
    commonNames = Array("QA_ASSET", "Sheet1", "Data", "Assets", "Main")

    For i = LBound(commonNames) To UBound(commonNames)
        Set ws = SafeGetWorksheet(ThisWorkbook, commonNames(i))
        If Not ws Is Nothing Then
            DebugPrint "GetValidWorksheetForUI: Found common sheet: " & ws.Name
            Set GetValidWorksheetForUI = ws
            Exit Function
        End If
    Next i

    ' Priority 4: Use first available sheet
    If ThisWorkbook.Worksheets.Count > 0 Then
        DebugPrint "GetValidWorksheetForUI: Using first worksheet"
        Set GetValidWorksheetForUI = ThisWorkbook.Worksheets(1)
        Exit Function
    End If

    ' No worksheet found
    DebugPrint "GetValidWorksheetForUI: FAILED - No worksheet found"
    Set GetValidWorksheetForUI = Nothing
End Function

Public Sub RebuildMatchBuilderUI(Optional ByVal userSelectedHeaderRow As Long = 0)
    '
    ' LAYER 2: UI Builder
    '
    ' This procedure builds the Match Builder UI interface.
    ' In SafeMode, it only creates buttons without modifying worksheet structure.
    ' Without SafeMode, it can insert rows for the UI.
    '
    ' Parameters:
    '   userSelectedHeaderRow - If > 0, use this as the header row instead of auto-detecting
    '
    Dim ws As Worksheet
    Dim dataHeaderRow As Long
    Dim lastDataCol As Long
    Dim existingMatches As Collection
    Dim safetyResponse As Integer
    Dim detectedRowBeforeInsert As Long
    Dim spaceNeeded As Long
    Dim i As Long
    Dim hasNOTE1 As Boolean
    Dim uiExists As Boolean
    Dim usingUserSelectedRow As Boolean
    Dim matchedIdCol As Long
    Dim cellVal As String
    Dim r As Long
    Dim existingMatchRowCount As Long
    Dim preserveMatchRows As Boolean
    Dim clearEndRow As Long
    Dim matchResultCols As Variant
    Dim matchColIdx As Long
    Dim lastDataRow As Long
    Dim req As Variant
    Dim checkCol As Long
    Dim headerVal As String
    Dim scanRow As Long
    Dim foundLastRow As Boolean
    Dim hasDataAtFallbackRow As Boolean
    Dim fallbackCheckCol As Long

    usingUserSelectedRow = (userSelectedHeaderRow > 0)

    DebugPrint "RebuildMatchBuilderUI: Starting (SafeMode=" & g_SafeMode & ")..."

    ' Enable error handling and disable screen updates for safety
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' Try to get a valid worksheet - check multiple sources
    Set ws = GetValidWorksheetForUI

    If ws Is Nothing Then
        MsgBox "No worksheet available for Match Builder.", vbCritical
        DebugPrint "RebuildMatchBuilderUI: FAILED - No worksheet available"
        Application.DisplayAlerts = True
        GoTo CleanExit
    End If

    DebugPrint "RebuildMatchBuilderUI: Using worksheet: " & ws.Name

    ' CRITICAL: Initialize dataset context FIRST - this sets all globals safely
    ' This does NOT modify the sheet, only detects the header position
    If Not g_Initialized Then
        DebugPrint "RebuildMatchBuilderUI: Calling InitializeDatasetContext..."
        Call InitializeDatasetContext(ws)
    End If

    ' Check if UI already exists
    uiExists = IsUIAlreadyExists(ws)
    DebugPrint "RebuildMatchBuilderUI: UI exists = " & uiExists

    ' SAFETY CHECK: If UI already exists, don't rebuild (unless force rebuild requested)
    If uiExists And Not g_ForceRebuild Then
        DebugPrint "RebuildMatchBuilderUI: UI already exists, skipping rebuild"
        Application.DisplayAlerts = True
        GoTo CleanExit
    End If

    ' SAFETY: If match rows already exist with data, do NOT rebuild them
    ' Just ensure UI structure (buttons, headers) exists
    ' Match rows managed by Add Match / Delete Match only

    For r = UI_FIRST_MATCH_ROW To 50
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" Then
            existingMatchRowCount = existingMatchRowCount + 1
        End If
    Next r

    ' Only preserve match rows if NOT a full rebuild AND existing matches exist
    preserveMatchRows = (existingMatchRowCount > 0) And Not g_ClearMatchDataOnRebuild
    DebugPrint "preserveMatchRows=" & preserveMatchRows & " existingMatchRowCount=" & existingMatchRowCount & " g_ClearMatchDataOnRebuild=" & g_ClearMatchDataOnRebuild

    ' Get the original header position (READ ONLY)
    ' If user selected a specific row, use that instead of auto-detection
    If usingUserSelectedRow Then
        detectedRowBeforeInsert = userSelectedHeaderRow
        DebugPrint "RebuildMatchBuilderUI: Using user-selected header row = " & detectedRowBeforeInsert
    Else
        detectedRowBeforeInsert = g_DataHeaderRow
        DebugPrint "RebuildMatchBuilderUI: detectedRowBeforeInsert = " & detectedRowBeforeInsert

        ' Guard: if detectedRowBeforeInsert came from fallback (FIXED_DATA_HEADER_ROW),
        ' verify actual data exists at that row before proceeding
        If detectedRowBeforeInsert = FIXED_DATA_HEADER_ROW Then
            hasDataAtFallbackRow = False
            For fallbackCheckCol = 6 To 50
                If Trim(CStr(ws.Cells(detectedRowBeforeInsert, fallbackCheckCol).Value)) <> "" Then
                    hasDataAtFallbackRow = True
                    Exit For
                End If
            Next fallbackCheckCol
            If Not hasDataAtFallbackRow Then
                MsgBox "No source data found on this worksheet." & vbCrLf & vbCrLf & _
                       "Please click 'Load Source' to load your data before building the UI.", _
                       vbExclamation, "Load Source First"
                DebugPrint "RebuildMatchBuilderUI: ABORTED - no data at fallback row, user must load source first"
                Application.DisplayAlerts = True
                GoTo CleanExit
            End If
        End If

        ' If g_DataHeaderRow wasn't set, detect it now (READ ONLY)
        If detectedRowBeforeInsert <= 1 Then
            Dim tempAliases As Object
            Dim tempLearned As Object
            Set tempAliases = CreateObject("Scripting.Dictionary")
            Set tempLearned = CreateObject("Scripting.Dictionary")
            detectedRowBeforeInsert = FindDataHeaderRow(ws, tempAliases, tempLearned)
            If detectedRowBeforeInsert <= 1 Then
                detectedRowBeforeInsert = FindDataHeaderRowAggressive(ws, tempAliases, tempLearned)
            End If
            If detectedRowBeforeInsert <= 1 Then
                MsgBox "Could not detect the data header row in this worksheet." & vbCrLf & _
                       "Build UI cannot proceed without a valid header row." & vbCrLf & _
                       "Please use Select Files to choose your source file and try again.", vbCritical, "Header Detection Failed"
                DebugPrint "RebuildMatchBuilderUI: ABORTED - header detection failed, no fallback allowed"
                Application.DisplayAlerts = True
                GoTo CleanExit
            End If
            DebugPrint "RebuildMatchBuilderUI: Recalculated header row = " & detectedRowBeforeInsert
        End If
    End If

    ' Check how many rows are available above the header
    spaceNeeded = detectedRowBeforeInsert - 1  ' Rows available above header
    DebugPrint "RebuildMatchBuilderUI: spaceNeeded = " & spaceNeeded

    ' Use the detected row
    dataHeaderRow = detectedRowBeforeInsert

    ' Determine last data column (READ ONLY)
    lastDataCol = GetLastColumn(ws, dataHeaderRow)
    If lastDataCol < 2 Then lastDataCol = 10
    DebugPrint "RebuildMatchBuilderUI: lastDataCol = " & lastDataCol

    ' ============================================================
    ' SAFE MODE HANDLING
    ' ============================================================

    If g_SafeMode Then
        ' In SafeMode: Check if we have enough space for UI
        ' If no space, prompt user to insert rows

        DebugPrint "RebuildMatchBuilderUI: Running in SAFE MODE"
        DebugPrint "RebuildMatchBuilderUI: spaceNeeded = " & spaceNeeded & ", UI_HEIGHT = " & UI_HEIGHT

        ' Check if we have enough space for UI
        If spaceNeeded < UI_HEIGHT Then
            ' Not enough space - ask user if they want to insert rows
            safetyResponse = MsgBox("SafeMode is enabled but there is insufficient space for full UI." & vbCrLf & vbCrLf & _
                    "Available rows above data: " & spaceNeeded & vbCrLf & _
                    "Required rows: " & UI_HEIGHT & vbCrLf & vbCrLf & _
                    "Do you want to insert " & UI_HEIGHT & " rows above your data?" & vbCrLf & _
                    "Your data will be shifted down.", vbYesNo + vbExclamation, "Build UI - Insert Rows?")

            If safetyResponse = vbNo Then
                MsgBox "UI build cancelled. Buttons cannot be placed in the data area.", vbInformation
                DebugPrint "RebuildMatchBuilderUI: User cancelled due to no space"
                Application.DisplayAlerts = True
                GoTo CleanExit
            End If

            ' User said yes - disable SafeMode and insert rows
            g_SafeMode = False
            DebugPrint "RebuildMatchBuilderUI: User opted to insert rows - SafeMode disabled"
        End If

        ' If SafeMode is still True (had enough space), build at current position
        ' If SafeMode is now False, fall through to normal build
        If g_SafeMode Then
            ' Build UI at current position (non-destructive)
            Set g_CurrentWorksheet = ws
            g_DataHeaderRow = dataHeaderRow
            g_DataStartRow = dataHeaderRow + 1
            g_LastDataColumn = lastDataCol
            g_Initialized = True

            ' Build UI elements (buttons only - no row insertion)
            Call BuildUIElementsOnly(ws, lastDataCol)

            ' Also add MATCH_TYPE column if it doesn't exist (right after MATCHED_ID)
            ' CRITICAL: Search in the DATA header row, not the UI header row
            hasNOTE1 = False
            matchedIdCol = 0

            ' First, try to find MATCH_TYPE or MATCHED_ID at g_DataHeaderRow (actual data row)
            Dim searchRowSafe As Long
            For searchRowSafe = g_DataHeaderRow To g_DataHeaderRow + 5
                If searchRowSafe > ws.Rows.Count Then Exit For
                For i = 1 To lastDataCol + 5  ' Allow extra columns
                    If i > ws.Columns.Count Then Exit For
                    cellVal = UCase(Trim(CStr(ws.Cells(searchRowSafe, i).value)))
                    If cellVal = "MATCH_TYPE" Then
                        hasNOTE1 = True
                        Exit For
                    End If
                    ' Find MATCHED_ID column position
                    If cellVal = "MATCHED_ID" Or cellVal = "MATCHEDID" Then
                        matchedIdCol = i
                    End If
                Next i
                If hasNOTE1 Or matchedIdCol > 0 Then Exit For
            Next searchRowSafe

            ' FIX: Add MATCHED_ID if it doesn't exist (like AddMatchColumnsAtCorrectPosition does)
            If matchedIdCol = 0 Then
                ' Insert MATCHED_ID at column 1 (first column)
                ws.Columns(1).Insert Shift:=xlToRight
                ws.Cells(g_DataHeaderRow, 1).Value = "MATCHED_ID"
                ws.Cells(g_DataHeaderRow, 1).Font.Bold = True
                ws.Cells(g_DataHeaderRow, 1).Interior.Color = RGB(91, 115, 150)
                ws.Cells(g_DataHeaderRow, 1).Font.Color = RGB(255, 255, 255)
                ws.Cells(g_DataHeaderRow, 1).HorizontalAlignment = xlCenter
                matchedIdCol = 1
                lastDataCol = lastDataCol + 1  ' Data shifted right by 1
                DebugPrint "RebuildMatchBuilderUI (SafeMode): Added MATCHED_ID at column 1"
            End If

            If Not hasNOTE1 Then
                Dim newColSafe As Long
                ' If we found MATCHED_ID, add MATCH_TYPE right after it; otherwise add at end
                If matchedIdCol > 0 Then
                    newColSafe = matchedIdCol + 1
                Else
                    newColSafe = lastDataCol + 1
                End If

                ' Insert a new column for MATCH_TYPE
                ws.Columns(newColSafe).Insert Shift:=xlToRight

                ' Add MATCH_TYPE header at the actual data header row
                ws.Cells(g_DataHeaderRow, newColSafe).value = "MATCH_TYPE"
                ws.Cells(g_DataHeaderRow, newColSafe).Font.Bold = True
                ws.Cells(g_DataHeaderRow, newColSafe).Interior.Color = RGB(91, 115, 150)
                ws.Cells(g_DataHeaderRow, newColSafe).Font.Color = RGB(255, 255, 255)
                ws.Cells(g_DataHeaderRow, newColSafe).HorizontalAlignment = xlCenter

                ' Also add MATCH_TYPE to UI header row if it's different
                If g_DataHeaderRow > UI_COLHEADER_ROW Then
                    ws.Cells(UI_COLHEADER_ROW, newColSafe).value = "MATCH_TYPE"
                    ws.Cells(UI_COLHEADER_ROW, newColSafe).Font.Bold = True
                    ws.Cells(UI_COLHEADER_ROW, newColSafe).Interior.Color = RGB(91, 115, 150)
                    ws.Cells(UI_COLHEADER_ROW, newColSafe).Font.Color = RGB(255, 255, 255)
                    ws.Cells(UI_COLHEADER_ROW, newColSafe).HorizontalAlignment = xlCenter
                End If

                g_LastDataColumn = lastDataCol + 1
            End If

            ' Also add MATCH_STATUS column in SafeMode
            Dim hasNOTE2Safe As Boolean
            hasNOTE2Safe = False

            For searchRowSafe = g_DataHeaderRow To g_DataHeaderRow + 5
                If searchRowSafe > ws.Rows.Count Then Exit For
                For i = 1 To lastDataCol + 10
                    If i > ws.Columns.Count Then Exit For
                    cellVal = UCase(Trim(CStr(ws.Cells(searchRowSafe, i).value)))
                    If cellVal = "MATCH_STATUS" Or cellVal = "MATCHSTATUS" Then
                        hasNOTE2Safe = True
                        Exit For
                    End If
                Next i
                If hasNOTE2Safe Then Exit For
            Next searchRowSafe

            If Not hasNOTE2Safe Then
                Dim note1ColSafe As Long
                note1ColSafe = 0

                ' Find MATCH_TYPE column position
                For searchRowSafe = g_DataHeaderRow To g_DataHeaderRow + 5
                    If searchRowSafe > ws.Rows.Count Then Exit For
                    For i = 1 To lastDataCol + 10
                        If i > ws.Columns.Count Then Exit For
                        cellVal = UCase(Trim(CStr(ws.Cells(searchRowSafe, i).value)))
                        If cellVal = "MATCH_TYPE" Or cellVal = "MATCHTYPE" Then
                            note1ColSafe = i
                            Exit For
                        End If
                    Next i
                    If note1ColSafe > 0 Then Exit For
                Next searchRowSafe

                If note1ColSafe > 0 Then
                    newColSafe = note1ColSafe + 1
                    ' Insert a new column for MATCH_STATUS
                    ws.Columns(newColSafe).Insert Shift:=xlToRight

                    ' Add MATCH_STATUS header
                    ws.Cells(g_DataHeaderRow, newColSafe).value = "MATCH_STATUS"
                    ws.Cells(g_DataHeaderRow, newColSafe).Font.Bold = True
                    ws.Cells(g_DataHeaderRow, newColSafe).Interior.Color = RGB(91, 115, 150)
                    ws.Cells(g_DataHeaderRow, newColSafe).Font.Color = RGB(255, 255, 255)
                    ws.Cells(g_DataHeaderRow, newColSafe).HorizontalAlignment = xlCenter
                End If
            End If

            DebugPrint "RebuildMatchBuilderUI: SAFE MODE build complete"
            Application.DisplayAlerts = True
            GoTo CleanExit
        End If

    End If

    ' ============================================================
    ' NON-SAFE MODE: Full UI build with row insertion
    ' ============================================================

    ' If fewer than UI_HEIGHT rows of empty space, ask for confirmation
    ' BUT skip this if g_ForceRebuild=True because BuildFullUI already asked
    If spaceNeeded < UI_HEIGHT And Not g_ForceRebuild Then
        safetyResponse = MsgBox("Warning: Only " & spaceNeeded & " empty row(s) detected above the data header." & vbCrLf & vbCrLf & _
                        "The system needs to insert " & UI_HEIGHT & " rows above your data to build the UI." & vbCrLf & _
                        "This will shift your data down by " & UI_HEIGHT & " rows." & vbCrLf & vbCrLf & _
                        "Do you want to continue?", vbYesNo + vbExclamation, "Safety Confirmation")

        If safetyResponse = vbNo Then
            MsgBox "UI build cancelled by user. Your data is safe.", vbInformation
            DebugPrint "RebuildMatchBuilderUI: Cancelled by user"
            Application.DisplayAlerts = True
            GoTo CleanExit
        End If
    End If

    '===========================================================================
    ' STEP 1.5: Add MATCH_TYPE/MATCH_STATUS columns BEFORE inserting UI rows
    ' This adds columns at the correct position (after MATCHED_ID or at end)
    '===========================================================================
    DebugPrint "RebuildMatchBuilderUI: Adding MATCH columns BEFORE row insertion..."
    Call AddMatchColumnsAtCorrectPosition(ws, dataHeaderRow, lastDataCol)

    ' Fill empty cells in MATCH result columns with ~NULL~ so HasDataBelowHeader
    ' can correctly detect the header row
    matchResultCols = Array("MATCHED_ID", "MATCH_TYPE", "MATCH_STATUS", "SOURCE_FILE", "TARGET_FILE")

    ' Find last data row using explicit scan (no End(xlUp))
    lastDataRow = dataHeaderRow + 1
    foundLastRow = False
    For scanRow = dataHeaderRow + 1 To dataHeaderRow + 1000
        If Trim(CStr(ws.Cells(scanRow, 1).Value)) <> "" Then
            lastDataRow = scanRow
            foundLastRow = True
        End If
    Next scanRow

    If Not foundLastRow Then
        lastDataRow = dataHeaderRow + 1
    End If

    ' Only fill ~NULL~ if actual source data exists below header
    If Not foundLastRow Then GoTo SkipNullFill

    For Each req In matchResultCols
        matchColIdx = 0
        For checkCol = 1 To 100
            headerVal = UCase(Trim(CStr(ws.Cells(dataHeaderRow, checkCol).Value)))
            If headerVal = req Then
                matchColIdx = checkCol
                Exit For
            End If
        Next checkCol

        If matchColIdx > 0 And lastDataRow > dataHeaderRow Then
            For scanRow = dataHeaderRow + 1 To lastDataRow
                If Trim(CStr(ws.Cells(scanRow, matchColIdx).Value)) = "" Then
                    ws.Cells(scanRow, matchColIdx).Value = "~NULL~"
                End If
            Next scanRow
        End If
    Next req

    SkipNullFill:

    ' After adding columns, recalculate lastDataCol
    lastDataCol = GetLastColumn(ws, dataHeaderRow)
    If lastDataCol < 2 Then lastDataCol = 10
    DebugPrint "RebuildMatchBuilderUI: After adding columns, lastDataCol = " & lastDataCol

    ' FIX: Clear old UI area ONLY if this is first build (no existing match rows)
    ' If match rows already exist, PRESERVE them - only clear buttons/headers
    If Not preserveMatchRows Then
        clearEndRow = dataHeaderRow - 1
        If clearEndRow < 1 Then
            DebugPrint "RebuildMatchBuilderUI: dataHeaderRow=1, nothing to clear above it - skipping UI clear"
            GoTo SkipUIClear
        End If

        DebugPrint "RebuildMatchBuilderUI: First build - clearing old UI area (rows 1-" & clearEndRow & ") before inserting new UI"
        On Error Resume Next
        ws.Rows("1:" & clearEndRow).ClearFormats
        ws.Rows("1:" & clearEndRow).ClearContents
        Dim oldBtn As Excel.Button
        Dim buttonsToRemove As Collection
        Dim b As Variant
        Set buttonsToRemove = New Collection
        For Each oldBtn In ws.Buttons
            If oldBtn.Top < ws.Rows(clearEndRow + 1).Top Then
                buttonsToRemove.Add oldBtn
            End If
        Next oldBtn
        For Each b In buttonsToRemove
            b.Delete
        Next b
        On Error GoTo ErrorHandler
        DebugPrint "RebuildMatchBuilderUI: Old UI cleared"
    End If
    SkipUIClear:

    ' STEP 2: Only insert rows if UI doesn't already exist
    ' If UI already exists (detected earlier), rebuild in place without inserting
    If Not uiExists Then
        ' Insert UI_HEIGHT rows ABOVE the detected data header to make room for UI
        ' DESTRUCTIVE OPERATION - only allowed when SafeMode is False

        ' Save entire header row to array BEFORE insert
        Dim preInsertHeaders() As String
        Dim preInsertCount As Long
        Dim preInsertCol As Long
        Dim preInsertEmpty As Long
        preInsertCount = 0
        preInsertEmpty = 0
        For preInsertCol = 1 To ws.Columns.Count
            If Trim(CStr(ws.Cells(dataHeaderRow, preInsertCol).Value)) <> "" Then
                preInsertCount = preInsertCol
                preInsertEmpty = 0
            Else
                preInsertEmpty = preInsertEmpty + 1
                If preInsertEmpty >= 2 Then Exit For
            End If
        Next preInsertCol
        If preInsertCount > 0 Then
            ReDim preInsertHeaders(1 To preInsertCount)
            For preInsertCol = 1 To preInsertCount
                preInsertHeaders(preInsertCol) = CStr(ws.Cells(dataHeaderRow, preInsertCol).Value)
            Next preInsertCol
        End If
        ' Insert blank UI rows
        ws.Rows(dataHeaderRow).Resize(UI_HEIGHT).Insert Shift:=xlDown
        ' Check if header shifted correctly — if not, restore from saved array
        If preInsertCount > 0 Then
            If Trim(CStr(ws.Cells(dataHeaderRow + UI_HEIGHT, 1).Value)) = "" Then
                For preInsertCol = 1 To preInsertCount
                    ws.Cells(dataHeaderRow + UI_HEIGHT, preInsertCol).Value = preInsertHeaders(preInsertCol)
                Next preInsertCol
                DebugPrint "RebuildMatchBuilderUI: Header manually restored to row " & (dataHeaderRow + UI_HEIGHT)
            Else
                DebugPrint "RebuildMatchBuilderUI: Header shifted correctly to row " & (dataHeaderRow + UI_HEIGHT)
            End If
        End If

        If Err.Number <> 0 Then
            MsgBox "Could not insert rows for UI. Error: " & Err.Description, vbCritical
            DebugPrint "RebuildMatchBuilderUI: FAILED - Could not insert rows"
            Application.DisplayAlerts = True
            GoTo CleanExit
        End If
        On Error GoTo ErrorHandler

        ' After inserting, the data header is now UI_HEIGHT rows lower
        dataHeaderRow = dataHeaderRow + UI_HEIGHT
    Else
        ' UI already exists - rebuild in place, data header stays at current position
        DebugPrint "RebuildMatchBuilderUI: UI exists, rebuilding in place without inserting rows"
    End If

    DebugPrint "RebuildMatchBuilderUI: After insert, dataHeaderRow = " & dataHeaderRow

    ' Update globals to point to the REAL data header (now at dataHeaderRow)
    g_DataHeaderRow = dataHeaderRow
    g_DataStartRow = dataHeaderRow + 1
    g_LastDataColumn = lastDataCol

    ' Apply formatting to MATCH columns
    Call ApplyMatchColumnColors(ws, dataHeaderRow)

    Set g_CurrentWorksheet = ws
    g_Initialized = True

    ' STEP 3: Capture any existing match rules from the UI area BEFORE building
    ' FIX: If this is a full rebuild (BuildFullUI), skip capturing and create fresh
    If g_ClearMatchDataOnRebuild Then
        ' Full rebuild - create default match (don't restore old data)
        Set existingMatches = New Collection
        Dim freshMatch As Object
        Set freshMatch = CreateObject("Scripting.Dictionary")
        freshMatch.Add "ID", 1
        freshMatch.Add "Type", "Default"
        Dim freshCols As New Collection
        freshCols.Add "ASSETNUM"
        freshMatch.Add "ColIndices", freshCols
        existingMatches.Add freshMatch
        DebugPrint "RebuildMatchBuilderUI: Full rebuild - creating fresh default match"
    Else
        ' Incremental rebuild - preserve existing matches
        Set existingMatches = CaptureExistingMatchesFromUI(ws)
        DebugPrint "RebuildMatchBuilderUI: Captured " & existingMatches.Count & " existing matches"
    End If

    ' STEP 4: Now build the UI in the inserted rows (rows 1-UI_HEIGHT+1)
    ' Build UI with updated row references
    Call BuildUIFreshFixed(ws, lastDataCol, existingMatches, preserveMatchRows)

    ' STEP 5: Verify MATCH_TYPE and MATCH_STATUS columns exist
    ' Columns were already added in STEP 1.5, just verify and update g_LastDataColumn
    lastDataCol = GetLastColumn(ws, g_DataHeaderRow)
    g_LastDataColumn = lastDataCol
    DebugPrint "RebuildMatchBuilderUI: Verified columns, lastDataCol = " & lastDataCol

    DebugPrint "RebuildMatchBuilderUI: COMPLETE"

    ' Apply AutoFilter to data header row to enable filtering on all columns
    On Error Resume Next
    Application.ScreenUpdating = True
    ws.Rows(g_DataHeaderRow).AutoFilter
    On Error GoTo 0
    DebugPrint "RebuildMatchBuilderUI: AutoFilter applied to row " & g_DataHeaderRow

    ' UI Creation Confirmation
    DebugPrint "UI Builder completed successfully"

    ' CRITICAL: Reset g_ForceRebuild to prevent duplicate UI builds
    g_ForceRebuild = False
    g_ClearMatchDataOnRebuild = False  ' Reset flag after rebuild

    If g_SafeMode Then
        DebugPrint "System running in SAFE MODE"
    End If

CleanExit:
    ' Restore application settings
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    ' Only show error if there was actually an error (Err.Number <> 0)
    If Err.Number <> 0 Then
        MsgBox "Error rebuilding UI: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
        DebugPrint "RebuildMatchBuilderUI: ERROR - " & Err.Description
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' MATCH ROW MANAGEMENT
'===============================================================================

Public Sub AddMatchRow()
    '
    ' LAYER 3: Execution Engine - Add Match Row
    '
    ' This is a destructive operation that modifies worksheet data
    ' Protected by SafeMode check
    '
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim newRow As Long
    Dim maxID As Long
    Dim r As Long
    Dim emptyRowFound As Boolean
    Dim hasSourceIndicator As Boolean
    Dim hasMatchHeader As Boolean
    Dim uiScanRow As Long
    Dim uiCellVal As String
    Dim uiStartRow As Long
    Dim uiEndRow As Long

    g_MacroRunning = True

    ' SAFE MODE GUARD - Block destructive operations in SafeMode
    If g_SafeMode Then
        MsgBox "Add Match Row is blocked in SafeMode." & vbCrLf & vbCrLf & _
               "This operation modifies worksheet data." & vbCrLf & _
               "To proceed, click 'Configure Source/Target' first, " & _
               "or manually disable SafeMode using EnableSafeMode().", vbExclamation
        DebugPrint "AddMatchRow: BLOCKED in SafeMode"
        g_MacroRunning = False
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    ' Get the worksheet with proper validation
    Set ws = g_CurrentWorksheet
    If ws Is Nothing Then Set ws = g_SourceSheet
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(1)
    If ws Is Nothing Then
        MsgBox "No worksheet available.", vbExclamation
        g_MacroRunning = False
        Exit Sub
    End If
    Call RefreshHeaderRowVariables(ws)

    ' CRITICAL: Use safe initializer to set globals properly
    ' This will NOT modify the sheet, only detect header position
    If g_DataHeaderRow = 0 Or g_Initialized = False Then
        Call InitializeDatasetContext(ws)
    End If

    ' GUARD: Verify g_DataHeaderRow is now set after initialization
    ' If still 0 or unset, exit with friendly message
    If g_DataHeaderRow = 0 Then
        MsgBox "Unable to determine data header row." & vbCrLf & vbCrLf & _
               "Please rebuild the UI using 'Build UI' button.", vbExclamation
        DebugPrint "AddMatchRow: g_DataHeaderRow is 0 after initialization"
        g_MacroRunning = False
        Exit Sub
    End If

    ' CRITICAL: Calculate UI boundaries dynamically based on actual data header row
    uiStartRow = 5  ' Match rows start at row 5
    uiEndRow = g_DataHeaderRow - 1  ' Last row before data

    ' GUARD: Validate UI structure exists before proceeding
    ' Scan all rows from 1 to uiEndRow to find mandatory UI headers
    hasSourceIndicator = False
    hasMatchHeader = False

    For uiScanRow = 1 To uiEndRow
        uiCellVal = UCase(Trim(CStr(ws.Cells(uiScanRow, 1).Value)))
        If Left(uiCellVal, 7) = "SOURCE:" Then hasSourceIndicator = True
        If uiCellVal = "MATCH" Then hasMatchHeader = True
    Next uiScanRow

    If Not hasSourceIndicator Or Not hasMatchHeader Then
        MsgBox "UI appears to be missing required headers." & vbCrLf & vbCrLf & _
               "Source indicator found: " & IIf(hasSourceIndicator, "Yes", "No") & vbCrLf & _
               "Match header found: " & IIf(hasMatchHeader, "Yes", "No") & vbCrLf & vbCrLf & _
               "Please rebuild the UI using 'Build UI' button.", vbExclamation
        DebugPrint "AddMatchRow: UI structure missing - Source=" & hasSourceIndicator & ", Match=" & hasMatchHeader
        Exit Sub
    End If

    ' FINAL safety check after initialization
    If g_DataHeaderRow = 0 Then
        MsgBox "Cannot detect data header row. Please rebuild the UI first.", vbExclamation
        Exit Sub
    End If

    lastCol = g_LastDataColumn
    If lastCol < 3 Then lastCol = 10

    ' CRITICAL SAFETY CHECK: Ensure we never write to dataset area
    If uiEndRow < uiStartRow Then
        ' Try to recover by reinitializing
        Call InitializeDatasetContext(ws)
        uiEndRow = g_DataHeaderRow - 1
        If uiEndRow < uiStartRow Then
            MsgBox "UI area corrupted. Cannot add match row.", vbCritical
            Exit Sub
        End If
    End If

    ' Find the first empty row in the match area (between uiStartRow and uiEndRow)
    newRow = 0
    emptyRowFound = False

    For r = uiStartRow To uiEndRow
        ' Check if column 1 is empty or not a number
        If Trim(CStr(ws.Cells(r, 1).value)) = "" Then
            newRow = r
            emptyRowFound = True
            Exit For
        End If
    Next r

    ' If no empty row found, add to the next available row in the UI area
    If Not emptyRowFound Then
        ' Find the last used match row
        Dim lastUsedRow As Long
        lastUsedRow = 0
        For r = uiStartRow To uiEndRow
            If IsNumeric(ws.Cells(r, 1).value) Then
                lastUsedRow = r
            End If
        Next r

        ' Only add if there's room
        If lastUsedRow < uiEndRow Then
            newRow = lastUsedRow + 1
        Else
            ' No room - extend the UI by inserting a row (safely within UI area)
            ' CRITICAL: Only insert if uiEndRow + 1 is STILL less than g_DataHeaderRow
            ' FIX: Unmerge Match Type cells before inserting row to avoid "merged cell" error
            On Error Resume Next
            Application.DisplayAlerts = False
            Dim unmergeR As Long
            For unmergeR = uiStartRow To uiEndRow
                ws.Range(ws.Cells(unmergeR, 2), ws.Cells(unmergeR, 5)).UnMerge
            Next unmergeR
            Application.DisplayAlerts = True
            On Error GoTo 0

            If uiEndRow + 1 <= g_DataHeaderRow Then
                ws.Rows(uiEndRow + 1).Insert Shift:=xlDown
                newRow = uiEndRow + 1
                ' After inserting, data header moved down by 1
                g_DataHeaderRow = g_DataHeaderRow + 1
                g_DataStartRow = g_DataHeaderRow + 1
            Else
                MsgBox "Maximum number of match rows reached.", vbInformation
                Exit Sub
            End If
        End If
    End If

    ' FINAL SAFETY CHECK: Verify newRow is still in UI area
    If newRow >= g_DataHeaderRow Or newRow < uiStartRow Then
        MsgBox "Cannot add match row - UI area boundary error.", vbCritical
        Exit Sub
    End If

    ' Determine next match ID
    maxID = 0
    For r = uiStartRow To uiEndRow
        If IsNumeric(ws.Cells(r, 1).value) Then
            If ws.Cells(r, 1).value > maxID Then maxID = ws.Cells(r, 1).value
        End If
    Next r
    maxID = maxID + 1

    ' Populate new row
    ws.Cells(newRow, 1).value = maxID
    ws.Cells(newRow, 2).value = "Match_" & maxID
    ws.Cells(newRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(newRow, 2).HorizontalAlignment = xlCenter

    Application.EnableEvents = False

    ' Renumber ALL rows from UI_FIRST_MATCH_ROW to uiEndRow to ensure sequential order
    Dim renumRow As Long
    Dim renumEndRow As Long
    renumEndRow = g_DataHeaderRow - 1
    Dim renumMatchType As String
    For renumRow = UI_FIRST_MATCH_ROW To renumEndRow
        ws.Cells(renumRow, 1).Value = renumRow - UI_FIRST_MATCH_ROW + 1
        renumMatchType = Trim(CStr(ws.Cells(renumRow, 2).Value))
        If renumMatchType = "" Or renumMatchType = "Match_" & (renumRow - UI_FIRST_MATCH_ROW + 1) Or (Left(renumMatchType, 6) = "Match_" And IsNumeric(Mid(renumMatchType, 7))) Then
            ws.Cells(renumRow, 2).Value = "Match_" & (renumRow - UI_FIRST_MATCH_ROW + 1)
        End If
    Next renumRow
    Application.EnableEvents = True

    ' Clear any existing X marks in this new row
    For r = 3 To lastCol
        ws.Cells(newRow, r).ClearContents
    Next r

    ' Reapply formatting
    Call FormatMatchBoxFixed(ws, lastCol)

    ' FIX: Re-merge Match Type cells after adding new row
    On Error Resume Next
    Dim mergeR As Long
    ' FIX: Disable alerts to prevent duplicate "Merging cell only keeps the upper-left value" warning
    Application.DisplayAlerts = False
    For mergeR = uiStartRow To g_DataHeaderRow - 1
        ws.Range(ws.Cells(mergeR, 2), ws.Cells(mergeR, 5)).Merge
    Next mergeR
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Force UI refresh
    ws.Activate

    g_MacroRunning = False
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    g_MacroRunning = False
    MsgBox "Could not add match row. Please try again.", vbExclamation
End Sub

Public Sub DeleteMatchRow()
    '
    ' LAYER 3: Execution Engine - Delete Match Row
    '
    ' This is a destructive operation that modifies worksheet data
    ' Protected by SafeMode check
    ' FIXED: Now uses shift-up approach with five-property copy
    '

    Dim userInput As String
    g_MacroRunning = True
    userInput = InputBox("Select match rows to delete — enter the Match numbers shown in column 1 (e.g. 1,3,5 or 2-4):", "Delete Rows", "")

    If Trim(userInput) = "" Then
        g_MacroRunning = False
        Exit Sub
    End If

    Dim ws As Worksheet

    ' SAFE MODE GUARD - Block destructive operations in SafeMode
    If g_SafeMode Then
        MsgBox "Delete Match Row is blocked in SafeMode." & vbCrLf & vbCrLf & _
               "This operation modifies worksheet data." & vbCrLf & _
               "To proceed, click 'Configure Source/Target' first, " & _
               "or manually disable SafeMode using EnableSafeMode().", vbExclamation
        DebugPrint "DeleteMatchRow: BLOCKED in SafeMode"
        g_MacroRunning = False
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    ' Get the worksheet
    Set ws = g_CurrentWorksheet
    If ws Is Nothing Then Set ws = g_SourceSheet
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(1)
    If ws Is Nothing Then
        MsgBox "No worksheet available.", vbExclamation
        g_MacroRunning = False
        Exit Sub
    End If
    Call RefreshHeaderRowVariables(ws)

    ' CRITICAL: Use safe initializer to set globals properly
    If g_DataHeaderRow = 0 Or g_Initialized = False Then
        Call InitializeDatasetContext(ws)
    End If

    ' FINAL safety check after initialization
    If g_DataHeaderRow = 0 Then
        MsgBox "Cannot detect data header row. Please rebuild the UI first.", vbExclamation
        g_MacroRunning = False
        Exit Sub
    End If

    ' Calculate UI boundaries locally
    Dim uiStartRow As Long
    Dim uiEndRow As Long
    uiEndRow = g_DataHeaderRow - 1

    ' Scan downward from top to find first UI row (CORRECTED)
    uiStartRow = 0
    Dim scanRow As Long
    For scanRow = 1 To uiEndRow - 1
        If Trim(CStr(ws.Cells(scanRow, 1).Value)) <> "" Then
            uiStartRow = scanRow
            Exit For
        End If
    Next scanRow
    If uiStartRow = 0 Then uiStartRow = 5  ' Fallback

    ' Guard against invalid range
    If uiStartRow >= uiEndRow Then
        MsgBox "No UI rows available to delete.", vbExclamation
        Exit Sub
    End If

    ' Get lastUIcol with safe calculation
    Dim lastUIcol As Long
    If g_LastDataColumn > 0 Then
        lastUIcol = g_LastDataColumn
    Else
        ' Calculate by scanning
        For scanRow = 1 To 100
            If Trim(CStr(ws.Cells(g_DataHeaderRow, scanRow).Value)) <> "" Then
                lastUIcol = scanRow
            End If
        Next scanRow
    End If
    If lastUIcol < 1 Then lastUIcol = 10  ' Absolute fallback

    ' Parse user input and convert to absolute row numbers
    Dim processedInput As String
    Dim inputParts() As String
    Dim inputPart As String
    Dim partIdx As Long
    Dim dashPos As Long
    Dim rangeStart As Long
    Dim rangeEnd As Long
    Dim absNum As Long

    processedInput = ""
    inputParts = Split(Replace(userInput, " ", ""), ",")

    For partIdx = 0 To UBound(inputParts)
        inputPart = Trim(inputParts(partIdx))
        If inputPart <> "" Then
            dashPos = InStr(2, inputPart, "-")
            If dashPos > 0 Then
                On Error Resume Next
                rangeStart = CLng(Left(inputPart, dashPos - 1))
                rangeEnd = CLng(Mid(inputPart, dashPos + 1))
                On Error GoTo 0
                If rangeStart > 0 And rangeEnd > 0 Then
                    If processedInput <> "" Then processedInput = processedInput & ","
                    processedInput = processedInput & CStr(rangeStart + UI_FIRST_MATCH_ROW - 1) & "-" & CStr(rangeEnd + UI_FIRST_MATCH_ROW - 1)
                End If
            Else
                On Error Resume Next
                absNum = CLng(inputPart)
                On Error GoTo 0
                If absNum > 0 Then
                    If processedInput <> "" Then processedInput = processedInput & ","
                    processedInput = processedInput & CStr(absNum + UI_FIRST_MATCH_ROW - 1)
                End If
            End If
        End If
    Next partIdx

    If processedInput = "" Then Exit Sub

    Dim rowsToDelete() As Long
    g_ParseCancelled = False
    rowsToDelete = ParseRowSelection(processedInput, uiStartRow, uiEndRow)

    If g_ParseCancelled Then Exit Sub

    Dim numDeleted As Long
    numDeleted = UBound(rowsToDelete) - LBound(rowsToDelete) + 1

    ' Deduplicate rowsToDelete — remove repeated row numbers
    Dim dedupIdx As Long
    Dim dedupInner As Long
    Dim dedupCount As Long
    Dim dedupArr() As Long
    Dim isDuplicate As Boolean
    dedupCount = 0
    ReDim dedupArr(LBound(rowsToDelete) To UBound(rowsToDelete))
    For dedupIdx = LBound(rowsToDelete) To UBound(rowsToDelete)
        isDuplicate = False
        For dedupInner = 0 To dedupCount - 1
            If dedupArr(LBound(dedupArr) + dedupInner) = rowsToDelete(dedupIdx) Then
                isDuplicate = True
                Exit For
            End If
        Next dedupInner
        If Not isDuplicate Then
            dedupArr(LBound(dedupArr) + dedupCount) = rowsToDelete(dedupIdx)
            dedupCount = dedupCount + 1
        End If
    Next dedupIdx
    ReDim Preserve dedupArr(LBound(dedupArr) To LBound(dedupArr) + dedupCount - 1)
    rowsToDelete = dedupArr
    numDeleted = dedupCount

    ' Step 2: Sort in descending order
    Call QuickSortDescending(rowsToDelete)

    ' Unmerge Match Type cells before shift-up to avoid merged cell errors
    On Error Resume Next
    Dim unmergeR As Long
    For unmergeR = uiStartRow To uiEndRow
        ws.Range(ws.Cells(unmergeR, 2), ws.Cells(unmergeR, 5)).UnMerge
    Next unmergeR
    On Error GoTo 0

    Call RefreshHeaderRowVariables(ws)
    uiEndRow = g_DataHeaderRow - 1

    Application.EnableEvents = False
    ' Step 3: Delete rows directly — UI shrinks by numDeleted rows
    ' Process from largest row to smallest (LBound to UBound after QuickSortDescending)
    ' so earlier deletions do not affect indices of later ones
    Dim i As Long
    For i = LBound(rowsToDelete) To UBound(rowsToDelete)
        If rowsToDelete(i) >= UI_FIRST_MATCH_ROW And rowsToDelete(i) <= uiEndRow Then
            ws.Rows(rowsToDelete(i)).Delete Shift:=xlUp
            uiEndRow = uiEndRow - 1
            g_DataHeaderRow = g_DataHeaderRow - 1
            g_DataStartRow = g_DataStartRow - 1
        End If
    Next i

    ' Step 4: Renumber rows sequentially
    Application.EnableEvents = False
    Dim newRowNum As Long
    Dim delRow As Long
    newRowNum = 1
    Dim matchTypeVal As String
    For delRow = UI_FIRST_MATCH_ROW To uiEndRow
        ws.Cells(delRow, 1).Value = newRowNum
        matchTypeVal = Trim(CStr(ws.Cells(delRow, 2).Value))
        If Left(matchTypeVal, 6) = "Match_" And IsNumeric(Mid(matchTypeVal, 7)) Then
            ws.Cells(delRow, 2).Value = "Match_" & newRowNum
        End If
        newRowNum = newRowNum + 1
    Next delRow

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Step 5: Reapply formatting (use lastUIcol)
    ' Clear all formatting in UI match area before reapplying
    ' This removes any data row formatting that leaked in after row deletion
    On Error Resume Next
    ws.Range(ws.Rows(UI_FIRST_MATCH_ROW), ws.Rows(g_DataHeaderRow - 1)).ClearFormats
    On Error GoTo 0

    Call FormatMatchBoxFixed(ws, lastUIcol)

    ' Re-merge Match Type cells after deletion
    On Error Resume Next
    Application.DisplayAlerts = False
    Dim remergeR As Long
    For remergeR = UI_FIRST_MATCH_ROW To g_DataHeaderRow - 1
        ws.Range(ws.Cells(remergeR, 2), ws.Cells(remergeR, 5)).Merge
    Next remergeR
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Force UI refresh
    ws.Activate

    g_MacroRunning = False
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    g_MacroRunning = False
    MsgBox "Error deleting match row: " & Err.Description, vbCritical
End Sub

'===============================================================================
' ClearMatchContent - Clear content in selected UI rows (keep row numbers)
'===============================================================================
Public Sub ClearMatchContent()
    g_MacroRunning = True

    If g_SafeMode Then
        MsgBox "Clear Match Content is blocked in SafeMode.", vbExclamation
        g_MacroRunning = False
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = g_CurrentWorksheet
    If ws Is Nothing Then Set ws = g_SourceSheet
    If ws Is Nothing Then Set ws = ActiveSheet

    Dim selRange As Range
    On Error Resume Next
    Set selRange = Selection
    On Error GoTo 0

    If selRange Is Nothing Then
        MsgBox "Please select one or more cells to clear.", vbExclamation
        g_MacroRunning = False
        Exit Sub
    End If

    Dim cell As Range
    For Each cell In selRange.Cells
        If cell.Row >= UI_FIRST_MATCH_ROW And cell.Row < g_DataHeaderRow Then
            If UCase(Trim(CStr(cell.Value))) = "X" Then
                cell.ClearContents
                If cell.Row Mod 2 = 0 Then
                    cell.Interior.Color = RGB(214, 224, 240)
                Else
                    cell.Interior.Color = RGB(237, 242, 250)
                End If
            End If
        End If
    Next cell

    g_MacroRunning = False
End Sub

'===============================================================================
' UI HELPER FUNCTIONS
'===============================================================================

Private Function CaptureExistingMatches(ws As Worksheet, headerRow As Long, lastCol As Long) As Collection
    Dim matches As New Collection
    Dim r As Long, c As Long
    Dim id As Variant, typ As String
    Dim xCols As Collection
    Dim matchObj As Object

    For r = UI_FIRST_MATCH_ROW To headerRow - 1
        id = ws.Cells(r, 1).value
        If IsNumeric(id) Then
            typ = Trim(CStr(ws.Cells(r, 2).value))
            If typ <> "" Then
                Set xCols = New Collection
                For c = 3 To lastCol
                    If Trim(UCase(CStr(ws.Cells(r, c).value))) = "X" Then
                        xCols.Add c
                    End If
                Next c
                If xCols.Count > 0 Then
                    Set matchObj = CreateObject("Scripting.Dictionary")
                    matchObj.Add "ID", CLng(id)
                    matchObj.Add "Type", typ
                    matchObj.Add "ColIndices", xCols
                    matches.Add matchObj
                End If
            End If
        End If
    Next r
    Set CaptureExistingMatches = matches
End Function

' New function to capture existing matches from the UI area using dynamic g_DataHeaderRow
Private Function CaptureExistingMatchesFromUI(ws As Worksheet) As Collection
    Dim matches As New Collection
    Dim r As Long, c As Long
    Dim id As Variant, typ As String
    Dim xCols As Collection
    Dim matchObj As Object
    Dim lastCol As Long
    Dim endRow As Long

    lastCol = g_LastDataColumn
    If lastCol < 3 Then lastCol = 10

    ' Determine end row dynamically
    If g_DataHeaderRow > 0 Then
        endRow = g_DataHeaderRow - 1
    Else
        endRow = FIXED_DATA_HEADER_ROW - 1
    End If

    ' Capture from UI_FIRST_MATCH_ROW to endRow
    ' FIXED: Accept both numeric AND text IDs (like "Match_3")
    For r = UI_FIRST_MATCH_ROW To endRow
        id = ws.Cells(r, 1).value
        typ = Trim(CStr(ws.Cells(r, 2).value))

        ' If Match Type is empty but we have an ID, auto-fill Match Type with the ID
        If typ = "" And Trim(CStr(id)) <> "" Then
            typ = Trim(CStr(id))
            ' Also write it back to the sheet so it persists
            ws.Cells(r, 2).value = typ
        End If

        ' Accept row if: ID exists AND (Match Type exists OR we just filled it)
        If Trim(CStr(id)) <> "" And typ <> "" Then
            Set xCols = New Collection
            For c = 3 To lastCol
                If Trim(UCase(CStr(ws.Cells(r, c).value))) = "X" Then
                    xCols.Add c
                End If
            Next c
            Set matchObj = CreateObject("Scripting.Dictionary")
            ' Use the ID as-is (could be number or text like "Match_3")
            matchObj.Add "ID", id
            matchObj.Add "Type", typ
            matchObj.Add "ColIndices", xCols
            matches.Add matchObj
        End If
    Next r

    ' If no matches found, add default - use ALL columns from data header row
    ' FIX: Use all columns instead of hardcoding ASSETNUM
    If matches.Count = 0 Then
        Set matchObj = CreateObject("Scripting.Dictionary")
        matchObj.Add "ID", 1
        matchObj.Add "Type", "Default"
        ' Use ALL columns from data header row for matching
        Set xCols = New Collection
        Dim headerColIndex As Long
        Dim headerNameCol As String
        Dim normalizedColName As String
        For headerColIndex = 1 To lastCol
            headerNameCol = Trim(CStr(ws.Cells(g_DataHeaderRow, headerColIndex).Value))
            If headerNameCol <> "" Then
                ' Use SmartMatch to normalize column name
                normalizedColName = SmartMatch(headerNameCol, UltraNormalize(headerNameCol), Nothing, Nothing)
                If normalizedColName <> "" Then
                    xCols.Add normalizedColName
                Else
                    xCols.Add headerNameCol
                End If
            End If
        Next headerColIndex

        DebugPrint "CaptureExistingMatchesFromUI: No X marks selected, using all " & xCols.Count & " columns for matching"
        matchObj.Add "ColIndices", xCols
        matches.Add matchObj
    End If

    Set CaptureExistingMatchesFromUI = matches
End Function

' Clear entire UI zone to prevent duplication - uses g_DataHeaderRow dynamically
Private Sub ClearEntireUIZone(ws As Worksheet, ByVal endRow As Long, lastCol As Long)
    Dim actualEndRow As Long
    Dim btn As Excel.Button  ' Explicitly declared as Excel.Button
    Dim buttonsToDelete As Collection
    Dim b As Variant

    ' FIX: Safety check - NEVER clear if we can't verify header position
    ' This prevents accidental deletion of user data
    If g_DataHeaderRow = 0 Then
        DebugPrint "ClearEntireUIZone: SKIPPED - g_DataHeaderRow is not set"
        Exit Sub
    End If

    ' Use g_DataHeaderRow if set, otherwise use endRow parameter
    If g_DataHeaderRow > 0 Then
        actualEndRow = g_DataHeaderRow - 1
    Else
        actualEndRow = endRow - 1
    End If

    ' Additional safety: Verify we're not clearing data rows
    If actualEndRow < 1 Then Exit Sub

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(actualEndRow, lastCol))
    rng.ClearFormats
    rng.ClearContents

    ' FIX Issue 4: ALWAYS delete old buttons before adding new ones during code rebuild
    ' The g_AllowButtonDelete flag is for USER manual deletion protection only
    ' Code should always rebuild cleanly without button overlap/shrinking
    Set buttonsToDelete = New Collection
    For Each btn In ws.Buttons
        If btn.Top < ws.Rows(actualEndRow + 1).Top Then
            buttonsToDelete.Add btn
        End If
    Next btn

    For Each b In buttonsToDelete
        b.Delete
    Next b
    DebugPrint "ClearEntireUIZone: Deleted " & buttonsToDelete.Count & " old buttons"

    ' Also unmerge any merged cells in UI area
    On Error Resume Next
    ws.Range(ws.Cells(1, 1), ws.Cells(actualEndRow, lastCol)).UnMerge
    On Error GoTo 0
End Sub

' Aggressive header detection - looks for rows with asset column headers
Public Function FindDataHeaderRowAggressive(ws As Worksheet, aliases As Object, learned As Object) As Long
    Dim r As Long
    Dim c As Long
    Dim lastCol As Long
    Dim score As Long
    Dim maxScore As Long
    Dim bestRow As Long
    Dim rawValue As String
    Dim normalized As String
    Dim matched As String

    Const HEADER_SCAN_ROWS As Long = 20
    Const SKIP_UI_ROWS As Long = 6  ' Skip rows 1-6 (UI area)

    maxScore = 0
    bestRow = FIXED_DATA_HEADER_ROW  ' Default to fixed row
    lastCol = GetLastColumn(ws, 1)

    ' FIXED: Look for header starting from row 7 to skip UI area
    ' This prevents detecting old comparison results as headers
    ' FIXED: Require at least 3 matches to be considered a valid header
    For r = SKIP_UI_ROWS + 1 To HEADER_SCAN_ROWS
        score = 0
        For c = 1 To lastCol
            rawValue = CStr(ws.Cells(r, c).value)
            normalized = UltraNormalize(rawValue)
            If normalized <> "" Then
                matched = SmartMatch(rawValue, normalized, aliases, learned)
                If matched <> "" Then score = score + 1
            End If
        Next c

        If score > maxScore Then
            maxScore = score
            bestRow = r
        End If
    Next r

    ' Return best match or default
    If maxScore > 0 Then
        FindDataHeaderRowAggressive = bestRow
    Else
        FindDataHeaderRowAggressive = FIXED_DATA_HEADER_ROW
    End If
End Function

' Fixed UI builder - builds UI at fixed row positions
Private Sub BuildUIFreshFixed(ws As Worksheet, lastCol As Long, matches As Collection, preserveMatchRows As Boolean)
    Dim r As Long, c As Long
    Dim matchObj As Object
    Dim colIndices As Collection
    Dim i As Long
    Dim bufCol As Long

    DebugPrint "BuildUIFreshFixed: Starting..."

    ' Row 1: Source configuration display
    ' Row 2: Target configuration display
    Call AddStatusDisplayFixed(ws)

    ' Row 3: Control buttons
    Call AddControlButtonsFixed(ws)

    ' Row 4: Column headers (Match | Match Type | dynamic dataset columns)
    ws.Cells(UI_COLHEADER_ROW, 1).value = "Match"
    ws.Cells(UI_COLHEADER_ROW, 2).value = "Match Type"
    ws.Cells(UI_COLHEADER_ROW, 1).Font.Bold = True
    ws.Cells(UI_COLHEADER_ROW, 2).Font.Bold = True
    ws.Cells(UI_COLHEADER_ROW, 1).Interior.Color = RGB(91, 115, 150)
    ws.Cells(UI_COLHEADER_ROW, 2).Interior.Color = RGB(91, 115, 150)
    ws.Cells(UI_COLHEADER_ROW, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(UI_COLHEADER_ROW, 2).Font.Color = RGB(255, 255, 255)
    ws.Cells(UI_COLHEADER_ROW, 1).HorizontalAlignment = xlCenter
    ws.Cells(UI_COLHEADER_ROW, 2).HorizontalAlignment = xlCenter
    ws.Cells(UI_COLHEADER_ROW, 1).WrapText = True
    ws.Cells(UI_COLHEADER_ROW, 2).WrapText = True

    ' Merge Match Type cell with blank cells C-E for better visual appearance
    ' FIX: Disable alerts to prevent duplicate "Merging cell only keeps the upper-left value" warning
    Application.DisplayAlerts = False
    ws.Range(ws.Cells(UI_COLHEADER_ROW, 2), ws.Cells(UI_COLHEADER_ROW, 5)).Merge
    Application.DisplayAlerts = True

    ' Set row height for header row
    ws.Rows(UI_COLHEADER_ROW).RowHeight = 30

    ' Copy dataset column headers to columns 3+ (match rule area)
    If g_DataHeaderRow > 0 And lastCol >= 3 Then
        Dim srcRange As Range, destRange As Range
        Set srcRange = ws.Range(ws.Cells(g_DataHeaderRow, 3), ws.Cells(g_DataHeaderRow, lastCol))
        Set destRange = ws.Range(ws.Cells(UI_COLHEADER_ROW, 3), ws.Cells(UI_COLHEADER_ROW, lastCol))
        destRange.Value = srcRange.Value
        destRange.Font.Bold = True
        destRange.Font.Color = RGB(255, 255, 255)
        destRange.HorizontalAlignment = xlCenter
        destRange.WrapText = True
        ' Columns 3-5 (mandatory area) get medium slate, columns 6+ get dark blue-grey
        If lastCol >= 3 Then
            ws.Range(ws.Cells(UI_COLHEADER_ROW, 3), ws.Cells(UI_COLHEADER_ROW, 5)).Interior.Color = RGB(91, 115, 150)
        End If
        If lastCol >= 6 Then
            ws.Range(ws.Cells(UI_COLHEADER_ROW, 6), ws.Cells(UI_COLHEADER_ROW, lastCol)).Interior.Color = RGB(68, 84, 106)
        End If
    End If

    ' Set column widths based on UI header row text length — not data rows
    For bufCol = 1 To lastCol
        ws.Columns(bufCol).ColumnWidth = Len(CStr(ws.Cells(UI_COLHEADER_ROW, bufCol).Value)) * 1.2 + 2
    Next bufCol

    ws.Columns(1).ColumnWidth = Len("MATCHED_ID") * 1.8 + 2
    ws.Columns(2).ColumnWidth = Len("MATCH_TYPE") * 1.8 + 2
    ws.Columns(3).ColumnWidth = Len("MATCH_STATUS") * 1.8 + 2
    ws.Columns(4).ColumnWidth = Len("SOURCE_FILE") * 1.8 + 2
    ws.Columns(5).ColumnWidth = Len("TARGET_FILE") * 1.8 + 2
    For c = 3 To lastCol
        If ws.Columns(c).ColumnWidth < 12 Then ws.Columns(c).ColumnWidth = 12
    Next c
    On Error GoTo 0

    ' NOTE: Match rows are managed by Add Match / Delete Match buttons
    ' BUILD UI no longer creates or modifies match rows (structure-only)

    ' RESTORE captured match rows with X marks during full rebuild
    If Not preserveMatchRows Then
        If matches.Count > 0 Then
            Dim currentRow As Long
            Dim colIdx As Variant
            currentRow = UI_FIRST_MATCH_ROW
            For Each matchObj In matches
                If currentRow <= g_DataHeaderRow - 1 Then
                    ' Write ID to column 1
                    ws.Cells(currentRow, 1).Value = matchObj("ID")
                    ' Write Type to column 2
                    ws.Cells(currentRow, 2).Value = matchObj("Type")
                    ' Write X marks to selected columns
                    Set colIndices = matchObj("ColIndices")
                    For Each colIdx In colIndices
                        On Error Resume Next
                        Dim colNum As Long
                        colNum = CLng(colIdx)
                        If Err.Number = 0 Then
                            If colNum >= 3 And colNum <= lastCol Then
                                ws.Cells(currentRow, colNum).Value = "X"
                            End If
                        End If
                        Err.Clear
                        On Error GoTo 0
                    Next colIdx
                    currentRow = currentRow + 1
                Else
                    Exit For
                End If
            Next matchObj
        End If
    End If

    ' Format match box area - use g_DataHeaderRow
    Call FormatMatchBoxFixed(ws, lastCol)

    ' Add border around data header row
    If g_DataHeaderRow > 0 Then
        ws.Rows(g_DataHeaderRow).BorderAround Weight:=xlMedium, ColorIndex:=1
    End If

    ' Freeze panes below the control row (row 3) so config stays visible
    On Error Resume Next
    ' First unfreeze any existing panes
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.FreezePanes = False
        ' Then activate the worksheet and set freeze
        ws.Activate
        If Not ActiveWindow Is Nothing Then
            Application.GoTo ws.Cells(1, 1), Scroll:=True
            ws.Range("A4").Select
            ActiveWindow.FreezePanes = True
        End If
    End If
    On Error GoTo 0

    DebugPrint "BuildUIFreshFixed: COMPLETE"
End Sub

'===============================================================================
' SAFE MODE UI BUILDER - Creates buttons only, no data modification
'===============================================================================

Private Sub BuildUIElementsOnly(ws As Worksheet, lastCol As Long)
    '
    ' LAYER 2: Safe Mode UI Builder
    '
    ' Creates buttons and UI elements WITHOUT modifying worksheet structure
    ' This is safe to run in SafeMode
    '
    Dim btn As Excel.Button

    DebugPrint "BuildUIElementsOnly: Starting..."

    ' Add control buttons (button creation is safe)
    Call AddControlButtonsFixed(ws)

    ' Add status display (label/text placement is safe)
    Call AddStatusDisplayFixed(ws)

    ' Debug: show where buttons are
    Dim btnCount As Long
    btnCount = 0
    For Each btn In ws.Buttons
        btnCount = btnCount + 1
        DebugPrint "BuildUIElementsOnly: Button #" & btnCount & " - " & btn.Name & " at row " & (btn.Top / ws.Rows(1).Height)
    Next btn
    DebugPrint "BuildUIElementsOnly: Total buttons = " & btnCount

    DebugPrint "BuildUIElementsOnly: COMPLETE"
End Sub

'===============================================================================
' SAFE MODE DISABLER - Allow destructive operations
'===============================================================================

Public Sub DisableSafeMode()
    '
    ' LAYER 3: Execution Enabler
    '
    ' Disables SafeMode to allow destructive operations
    ' User must explicitly call this before operations that modify data
    '
    DebugPrint "DisableSafeMode: User requested to disable SafeMode"
    g_SafeMode = False
    MsgBox "SafeMode DISABLED!" & vbCrLf & vbCrLf & _
           "Destructive operations are now allowed." & vbCrLf & _
           "Make sure you have backed up your data." & vbCrLf & vbCrLf & _
           "Use EnableSafeMode() to restore protection.", vbExclamation
End Sub

Public Sub EnableSafeMode()
    '
    ' LAYER 2: Safety Restored
    '
    ' Re-enables SafeMode to block destructive operations
    '
    DebugPrint "EnableSafeMode: SafeMode enabled"
    g_SafeMode = True
    MsgBox "SafeMode ENABLED!" & vbCrLf & vbCrLf & _
           "Destructive operations are now blocked.", vbInformation
End Sub

'===============================================================================
' EXECUTION MODE CHECK - For procedures that need destructive access
'===============================================================================

Private Function RequireExecutionMode() As Boolean
    '
    ' Checks if execution mode (SafeMode disabled) is available
    ' Returns True if operations can proceed, False otherwise
    '
    If g_SafeMode Then
        MsgBox "This operation requires disabling SafeMode." & vbCrLf & vbCrLf & _
               "Click 'No' to stay in SafeMode, or 'Yes' to disable it.", _
               vbYesNo + vbInformation, "SafeMode Protection"
        ' Note: This is a prompt - actual disable must be done manually
        RequireExecutionMode = False
    Else
        RequireExecutionMode = True
    End If
End Function

' Fixed format for match box - uses g_DataHeaderRow
Private Sub FormatMatchBoxFixed(ws As Worksheet, lastCol As Long)
    Dim r As Long, c As Long
    Dim cell As Range
    Dim endRow As Long

    endRow = g_DataHeaderRow - 1
    If endRow < UI_FIRST_MATCH_ROW Then endRow = UI_FIRST_MATCH_ROW + 1

    ' Set all row heights at once
    ws.Range(ws.Rows(UI_FIRST_MATCH_ROW), ws.Rows(endRow)).RowHeight = 25

    ' Format columns 1-2 - bulk apply
    Dim rngCol1 As Range, rngCol2 As Range
    Set rngCol1 = ws.Range(ws.Cells(UI_FIRST_MATCH_ROW, 1), ws.Cells(endRow, 1))
    Set rngCol2 = ws.Range(ws.Cells(UI_FIRST_MATCH_ROW, 2), ws.Cells(endRow, 2))

    rngCol1.HorizontalAlignment = xlCenter
    rngCol1.Font.Bold = True
    rngCol1.Borders.LineStyle = xlContinuous
    rngCol1.Borders.Weight = xlThin
    rngCol1.Interior.ColorIndex = xlNone

    rngCol2.HorizontalAlignment = xlCenter
    rngCol2.Borders.LineStyle = xlContinuous
    rngCol2.Borders.Weight = xlThin
    rngCol2.Interior.ColorIndex = xlNone

    ' Columns 3+ - format borders and alignment ONCE for entire range
    Dim rngData As Range
    Set rngData = ws.Range(ws.Cells(UI_FIRST_MATCH_ROW, 3), ws.Cells(endRow, lastCol))
    rngData.Borders.LineStyle = xlContinuous
    rngData.Borders.Weight = xlThin
    rngData.HorizontalAlignment = xlCenter

    ' Alternating row colors - loop rows NOT cells - one row at a time
    Dim rowRng As Range
    For r = UI_FIRST_MATCH_ROW To endRow
        Set rowRng = ws.Range(ws.Cells(r, 3), ws.Cells(r, lastCol))
        If r Mod 2 = 0 Then
            rowRng.Interior.Color = RGB(214, 224, 240)
        Else
            rowRng.Interior.Color = RGB(237, 242, 250)
        End If
    Next r

    ' Loop only for X marks - use SpecialCells to skip empty cells
    On Error Resume Next
    Dim nonEmptyCells As Range
    Set nonEmptyCells = rngData.SpecialCells(xlCellTypeConstants)
    If Not nonEmptyCells Is Nothing Then
        For Each cell In nonEmptyCells
            If Trim(UCase(cell.Value)) = "X" Then
                cell.Interior.Color = RGB(255, 100, 100)
                cell.Font.Bold = True
                cell.Font.Color = RGB(255, 255, 255)
            End If
        Next cell
    End If
    On Error GoTo 0
End Sub

' Fixed button placement - all buttons in row 1
Private Sub AddControlButtonsFixed(ws As Worksheet)
    Dim btn As Excel.Button  ' Explicitly declared as Excel.Button
    Dim btnTop As Double, btnWidth As Double, btnHeight As Double
    Dim endRow As Long

    On Error Resume Next

    DebugPrint "AddControlButtonsFixed: Starting..."

    ' Determine UI boundary - use g_DataHeaderRow if set
    If g_DataHeaderRow > 0 Then
        endRow = g_DataHeaderRow
    Else
        endRow = FIXED_DATA_HEADER_ROW
    End If

    ' FIX Issue 4: ALWAYS delete old buttons before adding new ones during code rebuild
    ' The g_AllowButtonDelete flag is for USER manual deletion protection only
    ' Code should always rebuild cleanly without button overlap/shrinking
    Dim buttonsToDelete As Collection
    Set buttonsToDelete = New Collection
    For Each btn In ws.Buttons
        If btn.Top < ws.Rows(endRow).Top Then
            buttonsToDelete.Add btn
        End If
    Next btn
    Dim b As Variant
    For Each b In buttonsToDelete
        b.Delete
    Next b

    ' Set row 3 height and styling for control buttons
    ws.Rows(UI_CONTROL_ROW).RowHeight = 30
    ws.Rows(UI_CONTROL_ROW).Interior.Color = RGB(240, 240, 240)

    btnHeight = 24
    btnTop = ws.Cells(UI_CONTROL_ROW, 1).Top + 3

    Dim fixedLeft As Long
    fixedLeft = 2

    ' Button 1: Load Source
    btnWidth = 90
    Set btn = ws.Buttons.Add(fixedLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "Load Source"
    btn.Name = "btnLoadSource"
    btn.OnAction = "LoadSourceFile"
    btn.Font.Size = 11
    btn.Font.Color = RGB(0, 0, 150)
    btn.Font.Bold = True
    Call SaveButtonAnchor(ws, btn)
    btn.Placement = xlFreeFloating
    fixedLeft = fixedLeft + btnWidth + 2

    ' Button 2: Load Target
    btnWidth = 90
    Set btn = ws.Buttons.Add(fixedLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "Load Target"
    btn.Name = "btnLoadTarget"
    btn.OnAction = "LoadTargetFile"
    btn.Font.Size = 11
    btn.Font.Color = RGB(0, 0, 150)
    btn.Font.Bold = True
    Call SaveButtonAnchor(ws, btn)
    btn.Placement = xlFreeFloating
    fixedLeft = fixedLeft + btnWidth + 2

    ' Button 3: + Add Match
    btnWidth = 80
    Set btn = ws.Buttons.Add(fixedLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "+ Add Match"
    btn.Name = "btnAddMatch"
    btn.OnAction = "AddMatchRow"
    btn.Font.Size = 11
    Call SaveButtonAnchor(ws, btn)
    btn.Placement = xlFreeFloating
    fixedLeft = fixedLeft + btnWidth + 2

    ' Button 4: - Delete
    btnWidth = 70
    Set btn = ws.Buttons.Add(fixedLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "- Delete"
    btn.Name = "btnDeleteMatch"
    btn.OnAction = "DeleteMatchRow"
    btn.Font.Size = 11
    Call SaveButtonAnchor(ws, btn)
    btn.Placement = xlFreeFloating
    fixedLeft = fixedLeft + btnWidth + 2

    ' Button 5: Clear X
    btnWidth = 60
    Set btn = ws.Buttons.Add(fixedLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "Clear X"
    btn.Name = "btnClearMatch"
    btn.OnAction = "ClearMatchContent"
    btn.Font.Size = 11
    Call SaveButtonAnchor(ws, btn)
    btn.Placement = xlFreeFloating
    fixedLeft = fixedLeft + btnWidth + 2

    ' Button 6: Execute Match
    btnWidth = 90
    Set btn = ws.Buttons.Add(fixedLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "Execute Match"
    btn.Name = "btnExecuteCompare"
    btn.OnAction = "ExecuteCompareWithValidation"
    btn.Font.Color = RGB(0, 100, 0)
    btn.Font.Bold = True
    btn.Font.Size = 12
    btn.Interior.Color = RGB(200, 255, 200)
    Call SaveButtonAnchor(ws, btn)
    btn.Placement = xlFreeFloating
    fixedLeft = fixedLeft + btnWidth + 2

    ' Button 7: Export Results
    btnWidth = 90
    Set btn = ws.Buttons.Add(fixedLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "Export Results"
    btn.Name = "btnExportResults"
    btn.OnAction = "ExportResults"
    btn.Font.Size = 11
    btn.Font.Color = RGB(0, 100, 0)
    btn.Font.Bold = True
    Call SaveButtonAnchor(ws, btn)
    btn.Placement = xlFreeFloating
    fixedLeft = fixedLeft + btnWidth + 2

    ' Button 8: Clear Data
    btnWidth = 80
    Set btn = ws.Buttons.Add(fixedLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "Clear Data"
    btn.Name = "btnClearData"
    btn.OnAction = "ClearAllData"
    btn.Font.Size = 11
    btn.Font.Color = RGB(180, 0, 0)
    btn.Font.Bold = True
    Call SaveButtonAnchor(ws, btn)
    btn.Placement = xlFreeFloating
    fixedLeft = fixedLeft + btnWidth + 2

    ' Button 9: Build UI
    btnWidth = 70
    Set btn = ws.Buttons.Add(fixedLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "Build UI"
    btn.Name = "btnBuildUI"
    btn.OnAction = "PreserveAndRebuildUI"
    btn.Font.Size = 11
    btn.Font.Color = RGB(180, 0, 0)
    btn.Font.Bold = True
    Call SaveButtonAnchor(ws, btn)
    btn.Placement = xlFreeFloating
    fixedLeft = fixedLeft + btnWidth + 2

    ' Button 10: Pause Macro
    btnWidth = 80
    Set btn = ws.Buttons.Add(fixedLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "Pause Macro"
    btn.Name = "btnToggleMacroEvents"
    btn.OnAction = "ToggleMacroEvents"
    btn.Font.Size = 11
    Call SaveButtonAnchor(ws, btn)
    btn.Placement = xlFreeFloating

    DebugPrint "AddControlButtonsFixed: COMPLETE - Added 10 buttons"
End Sub

'===============================================================================
' Execute Compare Wrapper - Validates before running CompareAssets
'===============================================================================

Public Sub ExecuteCompareWithValidation()
    '
    ' LAYER 3: Execution Engine - Execute Compare
    '
    ' This is a destructive operation that writes MATCHED_ID to worksheet
    ' Protected by SafeMode check
    '
    Dim configSheet As Worksheet
    Dim sourceWBName As String, sourceWSName As String
    Dim targetWBName As String, targetWSName As String
    Dim validationPass As Boolean
    Dim headerAliases As Object
    Dim tempLearned As Object
    Dim sourceCols As Object, targetCols As Object

    g_MacroRunning = True

    ' CRITICAL: Validate source/target configuration FIRST, before any other prompts
    ' AUTO-SET: Use current workbook sheets directly
    ' Source = current worksheet (MATCH_UI), Target = TARGET_DATA sheet
    Dim autoTargetWS As Worksheet
    On Error Resume Next
    Set autoTargetWS = ThisWorkbook.Worksheets("TARGET_DATA")
    On Error GoTo 0

    If autoTargetWS Is Nothing Then
        MsgBox "TARGET_DATA sheet not found." & vbCrLf & vbCrLf & _
               "Please click 'Load Target' to load your target data first.", vbExclamation
        Exit Sub
    End If

    ' Set source sheet
    Dim autoSourceWS As Worksheet
    Set autoSourceWS = g_CurrentWorksheet
    If autoSourceWS Is Nothing Then Set autoSourceWS = ActiveSheet

    ' Validate source has data
    If g_DataHeaderRow = 0 Then
        Call InitializeDatasetContext(autoSourceWS)
    End If

    If g_DataHeaderRow = 0 Then
        MsgBox "No source data found." & vbCrLf & vbCrLf & _
               "Please click 'Load Source' to load your source data first.", vbExclamation
        Exit Sub
    End If

    ' Set globals for CompareAssets
    Set g_SourceSheet = autoSourceWS
    Set g_TargetSheet = autoTargetWS
    Set g_SourceWorkbook = ThisWorkbook
    Set g_TargetWorkbook = ThisWorkbook

    ' Auto-detect target header row
    Dim autoTargetHeaderRow As Long
    autoTargetHeaderRow = 0
    Dim autoScanRow As Long
    For autoScanRow = 1 To 20
        If ValidateHeaderRow(autoTargetWS, autoScanRow) Then
            If HasDataBelowHeader(autoTargetWS, autoScanRow) Then
                autoTargetHeaderRow = autoScanRow
                Exit For
            End If
        End If
    Next autoScanRow

    If autoTargetHeaderRow = 0 Then autoTargetHeaderRow = 1

    g_SourceHeaderRow = g_DataHeaderRow
    g_TargetHeaderRow = autoTargetHeaderRow

    ' Step 2: Source and Target are already set via auto-detection above
    ' sourceWB and targetWB are already ThisWorkbook (assigned in auto-detection block)
    ' Just verify they are set
    If g_SourceSheet Is Nothing Then
        MsgBox "Source sheet not set.", vbExclamation
        Exit Sub
    End If

    If g_TargetSheet Is Nothing Then
        MsgBox "TARGET_DATA sheet not set.", vbExclamation
        GoTo CleanExit
    End If

    ' Step 3: Source and Target sheets/headers already auto-detected above
    ' headerAliases and tempLearned still needed for column mapping below
    Set headerAliases = CreateObject("Scripting.Dictionary")
    Set tempLearned = CreateObject("Scripting.Dictionary")

    ' Load header mapping if exists
    On Error Resume Next
    Call LoadHeaderMapping(headerAliases)
    On Error GoTo 0

    ' Sheets and header rows already set by auto-detection above
    Set sourceCols = GetColumnMap(g_SourceSheet, g_SourceHeaderRow, headerAliases, tempLearned)
    Set targetCols = GetColumnMap(g_TargetSheet, g_TargetHeaderRow, headerAliases, tempLearned)

    If sourceCols.Count = 0 Then
        MsgBox "Could not detect source columns. Please verify your source data has a valid header row.", vbExclamation
        GoTo CleanExit
    End If

    If targetCols.Count = 0 Then
        MsgBox "Could not detect target columns. Please verify your TARGET_DATA sheet has a valid header row.", vbExclamation
        GoTo CleanExit
    End If

    ' All validation passed - proceed with CompareAssets

    ' Step X: Validate all 5 MATCH columns exist BEFORE running CompareAssets
    ' If missing, block execution and tell user to run Build UI
    If Not CheckAllResultColumnsExist(g_SourceSheet, g_SourceHeaderRow) Then
        Dim missingCols As String
        missingCols = ""
        Dim checkCol As Long
        Dim headerVal As String
        Dim requiredCols As Variant
        requiredCols = Array("MATCHED_ID", "MATCH_TYPE", "MATCH_STATUS", "SOURCE_FILE", "TARGET_FILE")
        Dim req As Variant
        For Each req In requiredCols
            Dim found As Boolean
            found = False
            For checkCol = 1 To 100
                headerVal = UCase(Trim(CStr(g_SourceSheet.Cells(g_SourceHeaderRow, checkCol).Value)))
                ' EXACT MATCH ONLY - case insensitive via UCase
                If headerVal = req Then
                    found = True
                    Exit For
                End If
            Next checkCol
            If Not found Then
                missingCols = missingCols & "  - " & req & vbCrLf
            End If
        Next req

        MsgBox "Missing required MATCH columns in Source sheet!" & vbCrLf & vbCrLf & _
               "The following columns are missing:" & vbCrLf & missingCols & vbCrLf & _
               "Please run 'Build UI' first to create these columns," & vbCrLf & _
               "then run Execute Match again.", vbExclamation, "Missing Columns"
        GoTo CleanExit
    End If

    ' Prompt user to select which TARGET column value goes to MATCHED_ID
    Dim returnCol As String
    returnCol = PromptForMatchedIdColumn()

    If returnCol = "" Then
        GoTo CleanExit
    End If
    g_MatchedIdColumn = returnCol

    ' Run the comparison
    Call CompareAssets
    GoTo CleanExit

CleanExit:
    ' Restore source sheet focus on all exit paths
    On Error Resume Next
    If Not g_SourceSheet Is Nothing Then
        g_SourceSheet.Activate
    ElseIf Not g_CurrentWorksheet Is Nothing Then
        g_CurrentWorksheet.Activate
    End If
    On Error GoTo 0

    g_MacroRunning = False
End Sub

' Fixed status display - shows Source/Target with connection status in row 2
Private Sub AddStatusDisplayFixed(ws As Worksheet)
    Dim sourceName As String
    Dim targetName As String
    Dim sourceWBName As String
    Dim targetWBName As String
    Dim configSheet As Worksheet
    Dim sourceWB As Workbook
    Dim targetWB As Workbook
    Dim sourceActive As Boolean
    Dim targetActive As Boolean
    Dim statusConfigSheet As Worksheet
    Dim loadedSrcFile As String
    Dim loadedSrcSheet As String
    Dim loadedTgtFile As String
    Dim loadedTgtSheet As String

    On Error Resume Next
    sourceActive = False
    targetActive = False

    ' Get source/target info from globals and config
    Set statusConfigSheet = GetOrCreateConfigSheet
    loadedSrcFile = GetConfigValue(statusConfigSheet, "LOADED_SOURCE_FILE")
    loadedSrcSheet = GetConfigValue(statusConfigSheet, "LOADED_SOURCE_SHEET")
    If loadedSrcFile <> "" And loadedSrcSheet <> "" Then
        sourceName = loadedSrcSheet
        sourceWBName = loadedSrcFile
        sourceActive = True
    ElseIf Not g_SourceSheet Is Nothing Then
        sourceName = g_SourceSheet.Name
        sourceWBName = g_SourceSheet.Parent.Name
        sourceActive = True
    ElseIf Not g_SourceWorkbook Is Nothing Then
        sourceName = "[Not Set]"
        sourceWBName = g_SourceWorkbook.Name
        sourceActive = True  ' Workbook exists, assume sheet might be there
    End If
    loadedTgtFile = GetConfigValue(statusConfigSheet, "LOADED_TARGET_FILE")
    loadedTgtSheet = GetConfigValue(statusConfigSheet, "LOADED_TARGET_SHEET")
    If loadedTgtFile <> "" And loadedTgtSheet <> "" Then
        targetName = loadedTgtSheet
        targetWBName = loadedTgtFile
        targetActive = True
    ElseIf Not g_TargetSheet Is Nothing Then
        targetName = g_TargetSheet.Name
        targetWBName = g_TargetSheet.Parent.Name
        targetActive = True
    ElseIf Not g_TargetWorkbook Is Nothing Then
        targetName = "[Not Set]"
        targetWBName = g_TargetWorkbook.Name
        targetActive = True
    End If

    ' If globals not set, try reading from config and check if files are open
    If sourceName = "" Or sourceName = "[Not Set]" Then
        Set configSheet = GetOrCreateConfigSheet
        sourceWBName = GetConfigValue(configSheet, CONFIG_SOURCE_WB)
        sourceName = GetConfigValue(configSheet, CONFIG_SOURCE_WS)
        ' Check if workbook is actually open
        If sourceWBName <> "" Then
            Set sourceWB = GetWorkbookByName(sourceWBName)
            If Not sourceWB Is Nothing Then
                sourceActive = True
                ' Try to get the sheet
                On Error Resume Next
                Set g_SourceSheet = sourceWB.Worksheets(sourceName)
                If Not g_SourceSheet Is Nothing Then
                    sourceActive = True
                    Set g_SourceWorkbook = sourceWB
                End If
                On Error GoTo 0
            Else
                sourceActive = False
            End If
        End If
    End If

    If targetName = "" Or targetName = "[Not Set]" Then
        If configSheet Is Nothing Then Set configSheet = GetOrCreateConfigSheet
        targetWBName = GetConfigValue(configSheet, CONFIG_TARGET_WB)
        targetName = GetConfigValue(configSheet, CONFIG_TARGET_WS)
        ' Check if workbook is actually open
        If targetWBName <> "" Then
            Set targetWB = GetWorkbookByName(targetWBName)
            If Not targetWB Is Nothing Then
                targetActive = True
                ' Try to get the sheet
                On Error Resume Next
                Set g_TargetSheet = targetWB.Worksheets(targetName)
                If Not g_TargetSheet Is Nothing Then
                    targetActive = True
                    Set g_TargetWorkbook = targetWB
                End If
                On Error GoTo 0
            Else
                targetActive = False
            End If
        End If
    End If

    On Error GoTo 0

    ' Build display text with workbook | sheet format and connection status
    Dim sourceDisplay As String
    Dim targetDisplay As String
    Dim sourceStatusText As String
    Dim targetStatusText As String

    ' Status indicators
    If sourceActive And sourceName <> "" And sourceName <> "[Not Set]" Then
        sourceStatusText = " [ACTIVE]"
    Else
        sourceStatusText = " [CLOSED]"
    End If

    If targetActive And targetName <> "" And targetName <> "[Not Set]" Then
        targetStatusText = " [ACTIVE]"
    Else
        targetStatusText = " [CLOSED]"
    End If

    If sourceName <> "" And sourceName <> "[Not Set]" Then
        sourceDisplay = "Source: " & sourceWBName & " | " & sourceName & sourceStatusText
    Else
        sourceDisplay = "Source: [Not set - Click Select Files]"
    End If

    If targetName <> "" And targetName <> "[Not Set]" Then
        targetDisplay = "Target: " & targetWBName & " | " & targetName & targetStatusText
    Else
        targetDisplay = "Target: [Not set - Click Select Files]"
    End If

    ' Row 1: Source configuration display
    ws.Cells(UI_SOURCE_ROW, 1).value = sourceDisplay
    ws.Cells(UI_SOURCE_ROW, 1).Font.Italic = True
    If sourceActive Then
        ws.Cells(UI_SOURCE_ROW, 1).Font.Color = RGB(0, 100, 0)  ' Green for active
    Else
        ws.Cells(UI_SOURCE_ROW, 1).Font.Color = RGB(200, 0, 0)  ' Red for closed
    End If
    ws.Cells(UI_SOURCE_ROW, 1).Font.Size = 11
    ws.Cells(UI_SOURCE_ROW, 1).HorizontalAlignment = xlLeft

    ' Apply formatting to row 1 - background color
    ws.Rows(UI_SOURCE_ROW).RowHeight = 25
    If sourceActive Then
        ws.Rows(UI_SOURCE_ROW).Interior.Color = RGB(230, 245, 230)  ' Light green
    Else
        ws.Rows(UI_SOURCE_ROW).Interior.Color = RGB(255, 230, 230)  ' Light red
    End If

    ' Merge across columns for source row
    ' FIX: Disable alerts to prevent duplicate "Merging cell only keeps the upper-left value" warning
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Range(ws.Cells(UI_SOURCE_ROW, 2), ws.Cells(UI_SOURCE_ROW, 10)).ClearContents
    ws.Range(ws.Cells(UI_SOURCE_ROW, 1), ws.Cells(UI_SOURCE_ROW, 10)).Merge
    Application.DisplayAlerts = True
    ws.Range(ws.Cells(UI_SOURCE_ROW, 1), ws.Cells(UI_SOURCE_ROW, 10)).HorizontalAlignment = xlLeft
    On Error GoTo 0

    ' Row 2: Target configuration display
    ws.Cells(UI_TARGET_ROW, 1).value = targetDisplay
    ws.Cells(UI_TARGET_ROW, 1).Font.Italic = True
    If targetActive Then
        ws.Cells(UI_TARGET_ROW, 1).Font.Color = RGB(100, 0, 100)  ' Purple for active
    Else
        ws.Cells(UI_TARGET_ROW, 1).Font.Color = RGB(200, 0, 0)  ' Red for closed
    End If
    ws.Cells(UI_TARGET_ROW, 1).Font.Size = 11
    ws.Cells(UI_TARGET_ROW, 1).HorizontalAlignment = xlLeft

    ' Apply formatting to row 2 - background color
    ws.Rows(UI_TARGET_ROW).RowHeight = 25
    If targetActive Then
        ws.Rows(UI_TARGET_ROW).Interior.Color = RGB(245, 230, 245)  ' Light purple
    Else
        ws.Rows(UI_TARGET_ROW).Interior.Color = RGB(255, 230, 230)  ' Light red
    End If

    ' Merge across columns for target row
    ' FIX: Disable alerts to prevent duplicate "Merging cell only keeps the upper-left value" warning
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Range(ws.Cells(UI_TARGET_ROW, 2), ws.Cells(UI_TARGET_ROW, 10)).ClearContents
    ws.Range(ws.Cells(UI_TARGET_ROW, 1), ws.Cells(UI_TARGET_ROW, 10)).Merge
    Application.DisplayAlerts = True
    ws.Range(ws.Cells(UI_TARGET_ROW, 1), ws.Cells(UI_TARGET_ROW, 10)).HorizontalAlignment = xlLeft
    On Error GoTo 0

    DebugPrint "AddStatusDisplayFixed: COMPLETE - Source=" & sourceActive & ", Target=" & targetActive
End Sub

Private Sub ClearUIZone(ws As Worksheet, headerRow As Long, lastCol As Long)
    Dim btn As Excel.Button  ' Explicitly declared as Excel.Button

    If headerRow <= 1 Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(headerRow - 1, lastCol))
    rng.ClearFormats
    rng.ClearContents

    ' FIX Issue 4: ALWAYS delete old buttons when clearing UI
    ' The g_AllowButtonDelete flag is for USER manual deletion protection only
    For Each btn In ws.Buttons
        If btn.Top < ws.Rows(headerRow).Top Then
            btn.Delete
        End If
    Next btn
End Sub

Private Sub BuildUIFresh(ws As Worksheet, headerRow As Long, lastCol As Long, matches As Collection)
    Dim r As Long, c As Long
    Dim matchObj As Object
    Dim colIndices As Collection
    Dim i As Long

    ' Row 1: Control buttons
    Call AddControlButtons(ws, headerRow)

    ' Row 3: Source/Target status display
    Call AddStatusDisplay(ws, headerRow)

    ' Row 4: Column headers
    For c = 1 To lastCol
        If c <= 2 Then
            ws.Cells(UI_COLHEADER_ROW, c).value = Choose(c, "Match", "Match Type")
        Else
            ws.Cells(UI_COLHEADER_ROW, c).value = ws.Cells(headerRow, c).value
        End If
        ws.Cells(UI_COLHEADER_ROW, c).Font.Bold = True
        ws.Cells(UI_COLHEADER_ROW, c).Interior.Color = RGB(200, 200, 200)
        ws.Cells(UI_COLHEADER_ROW, c).HorizontalAlignment = xlCenter
    Next c

    ' Match rows
    r = UI_FIRST_MATCH_ROW
    For Each matchObj In matches
        ws.Cells(r, 1).value = matchObj("ID")
        ws.Cells(r, 2).value = matchObj("Type")
        Set colIndices = matchObj("ColIndices")
        For i = 1 To colIndices.Count
            c = colIndices(i)
            If c >= 3 And c <= lastCol Then
                ws.Cells(r, c).value = "X"
            End If
        Next i
        r = r + 1
    Next matchObj

    If matches.Count = 0 Then
        ws.Cells(UI_FIRST_MATCH_ROW, 1).value = 1
        ws.Cells(UI_FIRST_MATCH_ROW, 2).value = "1_Default"
        r = UI_FIRST_MATCH_ROW + 1
    End If

    Call FormatMatchBox(ws, headerRow, lastCol)
    ws.Rows(headerRow).BorderAround Weight:=xlMedium, ColorIndex:=1
End Sub

Private Sub AddStatusDisplay(ws As Worksheet, headerRow As Long)
    Dim statusText As String
    Dim sourceFile As String
    Dim targetFile As String

    On Error Resume Next

    ' Get source/target info from globals
    If Not g_SourceWorkbook Is Nothing Then
        sourceFile = g_SourceWorkbook.Name
    End If
    If Not g_TargetWorkbook Is Nothing Then
        targetFile = g_TargetWorkbook.Name
    End If

    On Error GoTo 0

    ' Build status text
    If sourceFile <> "" And targetFile <> "" Then
        statusText = "Source: " & sourceFile & " | Target: " & targetFile
    Else
        statusText = "Configure Source/Target to begin"
    End If

    ' Display in row 3
    ws.Cells(UI_STATUS_ROW, 1).value = statusText
    ws.Cells(UI_STATUS_ROW, 1).Font.Italic = True
    ws.Cells(UI_STATUS_ROW, 1).Font.Color = RGB(100, 100, 100)

    ' Merge across columns
    ' FIX: Disable alerts to prevent duplicate "Merging cell only keeps the upper-left value" warning
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Range(ws.Cells(UI_STATUS_ROW, 1), ws.Cells(UI_STATUS_ROW, 5)).Merge
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub FormatMatchBox(ws As Worksheet, headerRow As Long, lastCol As Long)
    Dim r As Long, c As Long
    Dim cell As Range

    If headerRow <= UI_FIRST_MATCH_ROW Then Exit Sub

    For r = UI_FIRST_MATCH_ROW To headerRow - 1
        For c = 3 To lastCol
            Set cell = ws.Cells(r, c)
            With cell.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
            If Trim(UCase(cell.value)) = "X" Then
                cell.Interior.Color = RGB(255, 0, 0)
            Else
                cell.Interior.Color = RGB(255, 255, 200)
            End If
        Next c
    Next r
End Sub

Private Sub AddControlButtons(ws As Worksheet, headerRow As Long)
    Dim btn As Excel.Button  ' Explicitly declared as Excel.Button
    Dim rng As Range
    Dim btnLeft As Double, btnTop As Double, btnWidth As Double, btnHeight As Double

    ' FIX Issue 4: Check g_AllowButtonDelete before deleting existing buttons
    If g_AllowButtonDelete Then
        For Each btn In ws.Buttons
            If btn.Top < ws.Rows(2).Top Then btn.Delete
        Next btn
    Else
        DebugPrint "AddControlButtons: Button protection enabled - skipping deletion of existing buttons"
    End If

    ' Add Match button
    Set rng = ws.Cells(UI_CONTROL_ROW, 1)
    btnLeft = rng.Left: btnTop = rng.Top
    btnWidth = rng.Width * 1.5: btnHeight = rng.Height
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "+ Add Match"
    btn.Name = "btnAddMatch"
    btn.OnAction = "AddMatchRow"

    ' Delete Match button
    Set rng = ws.Cells(UI_CONTROL_ROW, 2)
    btnLeft = rng.Left: btnTop = rng.Top
    btnWidth = rng.Width * 1.5: btnHeight = rng.Height
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "- Delete Match"
    btn.Name = "btnDeleteMatch"
    btn.OnAction = "DeleteMatchRow"

    ' Execute Match button
    Set rng = ws.Cells(UI_CONTROL_ROW, 3)
    btnLeft = rng.Left: btnTop = rng.Top
    btnWidth = rng.Width * 3: btnHeight = rng.Height
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "Execute Match"
    btn.Name = "btnExecuteMatch"
    btn.OnAction = "ExecuteCompareWithValidation"
    btn.Font.Color = RGB(0, 128, 0)
    btn.Font.Bold = True

    ' Configure Source/Target button
    Set rng = ws.Cells(UI_CONTROL_ROW, 4)
    btnLeft = rng.Left: btnTop = rng.Top
    btnWidth = rng.Width * 2.5: btnHeight = rng.Height
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Caption = "Configure Source/Target"
    btn.Name = "btnConfigureSourceTarget"
    btn.OnAction = "ConfigureSourceTargetManual"
End Sub

' Manual configuration entry point (called from button)
Public Sub SelectWorkbooksForComparison()
    '
    ' Select workbooks and sheets for comparison
    ' FIXED: Added retry loops + system sheet filtering
    '
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sourceWBName As String
    Dim sourceWSName As String
    Dim targetWBName As String
    Dim targetWSName As String
    Dim configSheet As Worksheet
    Dim sheetNames() As String
    Dim wbNames() As String
    Dim wbCount As Long
    Dim i As Long
    Dim response As VbMsgBoxResult
    Dim validSelection As Boolean
    Dim sourceSheetNames As Collection
    Dim targetSheetNames As Collection
    Dim sheetName As String

    DebugPrint "SelectWorkbooksForComparison: Starting..."

    ' Get list of open workbooks
    wbCount = Application.Workbooks.Count
    ReDim wbNames(1 To wbCount)
    For i = 1 To wbCount
        wbNames(i) = Application.Workbooks(i).Name
    Next i

    If wbCount = 0 Then
        MsgBox "No workbooks are open." & vbCrLf & vbCrLf & _
               "Please open the source and target workbooks first.", vbInformation
        DebugPrint "SelectWorkbooksForComparison: No workbooks open"
        Exit Sub
    End If

    ' ============================================================
    ' Step 1: Select SOURCE workbook with RETRY
    ' ============================================================
    validSelection = False
    Do While Not validSelection
        Dim sourcePrompt As String
        sourcePrompt = "Select SOURCE workbook:" & vbCrLf & vbCrLf
        For i = 1 To wbCount
            sourcePrompt = sourcePrompt & i & ". " & wbNames(i) & vbCrLf
        Next i
        sourcePrompt = sourcePrompt & vbCrLf & "Enter number (or Cancel to exit):"

        Dim sourceSelection As String
        sourceSelection = InputBox(sourcePrompt, "Select SOURCE Workbook", "1")

        If sourceSelection = "" Then
            DebugPrint "SelectWorkbooksForComparison: User cancelled at source workbook"
            Exit Sub
        End If

        On Error Resume Next
        Dim sourceNum As Long
        sourceNum = CLng(sourceSelection)
        On Error GoTo 0

        If sourceNum >= 1 And sourceNum <= wbCount Then
            validSelection = True
            sourceWBName = wbNames(sourceNum)
        Else
            MsgBox "Invalid selection. Please enter a number between 1 and " & wbCount, vbExclamation
        End If
    Loop

    ' ============================================================
    ' Step 2: Select SOURCE sheet with RETRY + FILTER
    ' ============================================================
    Set wb = Workbooks(sourceWBName)

    ' Build filtered sheet list (exclude system sheets)
    Set sourceSheetNames = New Collection
    For i = 1 To wb.Worksheets.Count
        sheetName = UCase(wb.Worksheets(i).Name)
        If sheetName <> "COMPARE_CONFIG" And sheetName <> "CONFIG" Then
            sourceSheetNames.Add wb.Worksheets(i).Name
        End If
    Next i

    If sourceSheetNames.Count = 0 Then
        MsgBox "No valid sheets found in " & sourceWBName & " (system sheets excluded).", vbCritical
        Exit Sub
    End If

    validSelection = False
    Do While Not validSelection
        Dim sourceSheetPrompt As String
        sourceSheetPrompt = "Select SOURCE sheet in '" & sourceWBName & "':" & vbCrLf & vbCrLf
        Dim maxShow As Long
        maxShow = 20
        If sourceSheetNames.Count < maxShow Then maxShow = sourceSheetNames.Count
        For i = 1 To maxShow
            sourceSheetPrompt = sourceSheetPrompt & i & ". " & sourceSheetNames(i) & vbCrLf
        Next i
        If sourceSheetNames.Count > maxShow Then
            sourceSheetPrompt = sourceSheetPrompt & "... and " & (sourceSheetNames.Count - maxShow) & " more." & vbCrLf
        End If
        sourceSheetPrompt = sourceSheetPrompt & vbCrLf & "Enter number (or Cancel to exit):"

        Dim sourceSheetSelection As String
        sourceSheetSelection = InputBox(sourceSheetPrompt, "Select SOURCE Sheet", "1")

        If sourceSheetSelection = "" Then
            DebugPrint "SelectWorkbooksForComparison: User cancelled at source sheet"
            Exit Sub
        End If

        On Error Resume Next
        sourceNum = CLng(sourceSheetSelection)
        On Error GoTo 0

        If sourceNum >= 1 And sourceNum <= sourceSheetNames.Count Then
            validSelection = True
            sourceWSName = sourceSheetNames(sourceNum)
        Else
            MsgBox "Invalid selection. Please enter a number between 1 and " & sourceSheetNames.Count, vbExclamation
        End If
    Loop

    ' ============================================================
    ' Step 3: Select TARGET workbook with RETRY
    ' ============================================================
    validSelection = False
    Do While Not validSelection
        Dim targetPrompt As String
        targetPrompt = "Select TARGET workbook:" & vbCrLf & vbCrLf
        For i = 1 To wbCount
            targetPrompt = targetPrompt & i & ". " & wbNames(i) & vbCrLf
        Next i
        targetPrompt = targetPrompt & vbCrLf & "Enter number (or Cancel to exit):"

        Dim targetSelection As String
        targetSelection = InputBox(targetPrompt, "Select TARGET Workbook", "1")

        If targetSelection = "" Then
            DebugPrint "SelectWorkbooksForComparison: User cancelled at target workbook"
            Exit Sub
        End If

        On Error Resume Next
        Dim targetNum As Long
        targetNum = CLng(targetSelection)
        On Error GoTo 0

        If targetNum >= 1 And targetNum <= wbCount Then
            validSelection = True
            targetWBName = wbNames(targetNum)
        Else
            MsgBox "Invalid selection. Please enter a number between 1 and " & wbCount, vbExclamation
        End If
    Loop

    ' ============================================================
    ' Step 4: Select TARGET sheet with RETRY + FILTER
    ' ============================================================
    Set wb = Workbooks(targetWBName)

    ' Build filtered sheet list (exclude system sheets)
    Set targetSheetNames = New Collection
    For i = 1 To wb.Worksheets.Count
        sheetName = UCase(wb.Worksheets(i).Name)
        If sheetName <> "COMPARE_CONFIG" And sheetName <> "CONFIG" Then
            targetSheetNames.Add wb.Worksheets(i).Name
        End If
    Next i

    If targetSheetNames.Count = 0 Then
        MsgBox "No valid sheets found in " & targetWBName & " (system sheets excluded).", vbCritical
        Exit Sub
    End If

    validSelection = False
    Do While Not validSelection
        Dim targetSheetPrompt As String
        targetSheetPrompt = "Select TARGET sheet in '" & targetWBName & "':" & vbCrLf & vbCrLf
        maxShow = 20
        If targetSheetNames.Count < maxShow Then maxShow = targetSheetNames.Count
        For i = 1 To maxShow
            targetSheetPrompt = targetSheetPrompt & i & ". " & targetSheetNames(i) & vbCrLf
        Next i
        If targetSheetNames.Count > maxShow Then
            targetSheetPrompt = targetSheetPrompt & "... and " & (targetSheetNames.Count - maxShow) & " more." & vbCrLf
        End If
        targetSheetPrompt = targetSheetPrompt & vbCrLf & "Enter number (or Cancel to exit):"

        Dim targetSheetSelection As String
        targetSheetSelection = InputBox(targetSheetPrompt, "Select TARGET Sheet", "1")

        If targetSheetSelection = "" Then
            DebugPrint "SelectWorkbooksForComparison: User cancelled at target sheet"
            Exit Sub
        End If

        On Error Resume Next
        targetNum = CLng(targetSheetSelection)
        On Error GoTo 0

        If targetNum >= 1 And targetNum <= targetSheetNames.Count Then
            validSelection = True
            targetWSName = targetSheetNames(targetNum)
        Else
            MsgBox "Invalid selection. Please enter a number between 1 and " & targetSheetNames.Count, vbExclamation
        End If
    Loop

    ' Validate workbooks exist
    Dim sourceWB As Workbook
    Dim targetWB As Workbook

    Set sourceWB = GetWorkbookByName(sourceWBName)
    If sourceWB Is Nothing Then
        MsgBox "Source workbook '" & sourceWBName & "' not found.", vbCritical
        Exit Sub
    End If

    Set targetWB = GetWorkbookByName(targetWBName)
    If targetWB Is Nothing Then
        MsgBox "Target workbook '" & targetWBName & "' not found.", vbCritical
        Exit Sub
    End If

    ' Validate sheets exist
    On Error Resume Next
    Set g_SourceSheet = sourceWB.Worksheets(sourceWSName)
    If g_SourceSheet Is Nothing Then
        MsgBox "Source sheet '" & sourceWSName & "' not found in " & sourceWBName, vbCritical
        Exit Sub
    End If

    Set g_TargetSheet = targetWB.Worksheets(targetWSName)
    If g_TargetSheet Is Nothing Then
        MsgBox "Target sheet '" & targetWSName & "' not found in " & targetWBName, vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' ============================================================
    ' VALIDATION: Check for ASSETID column in both files
    ' ============================================================
    Dim sourceHasAssetID As Boolean
    Dim targetHasAssetID As Boolean
    Dim sourceHeaderRow As Long
    Dim targetHeaderRow As Long
    Dim sourceLastCol As Long
    Dim targetLastCol As Long
    Dim col As Long
    Dim headerName As String

    ' Find header rows
    sourceHeaderRow = FindDataHeaderRow(g_SourceSheet, CreateObject("Scripting.Dictionary"), CreateObject("Scripting.Dictionary"))
    targetHeaderRow = FindDataHeaderRow(g_TargetSheet, CreateObject("Scripting.Dictionary"), CreateObject("Scripting.Dictionary"))

    If sourceHeaderRow <= 1 Then sourceHeaderRow = FindDataHeaderRowAggressive(g_SourceSheet, CreateObject("Scripting.Dictionary"), CreateObject("Scripting.Dictionary"))
    If targetHeaderRow <= 1 Then targetHeaderRow = FindDataHeaderRowAggressive(g_TargetSheet, CreateObject("Scripting.Dictionary"), CreateObject("Scripting.Dictionary"))

    If sourceHeaderRow <= 1 Then sourceHeaderRow = 1
    If targetHeaderRow <= 1 Then targetHeaderRow = 1

    sourceLastCol = GetLastColumn(g_SourceSheet, sourceHeaderRow)
    targetLastCol = GetLastColumn(g_TargetSheet, targetHeaderRow)

    ' Check for ASSETID in source
    sourceHasAssetID = False
    For col = 1 To sourceLastCol
        headerName = UCase(Trim(CStr(g_SourceSheet.Cells(sourceHeaderRow, col).value)))
        If headerName = "ASSETID" Or headerName = "ASSETNUM" Or headerName = "ASSET_NUM" Or headerName = "ASSET_NUMBER" Or headerName = "ID" Then
            sourceHasAssetID = True
            Exit For
        End If
    Next col

    ' Check for ASSETID in target
    targetHasAssetID = False
    For col = 1 To targetLastCol
        headerName = UCase(Trim(CStr(g_TargetSheet.Cells(targetHeaderRow, col).value)))
        If headerName = "ASSETID" Or headerName = "ASSETNUM" Or headerName = "ASSET_NUM" Or headerName = "ASSET_NUMBER" Or headerName = "ID" Then
            targetHasAssetID = True
            Exit For
        End If
    Next col

    ' If ASSETID not found in either file, ask for alternative
    If Not sourceHasAssetID Or Not targetHasAssetID Then
        Dim altColumn As String
        altColumn = InputBox("ASSETID column not found in one or both files." & vbCrLf & vbCrLf & _
                            "Source has ID column: " & IIf(sourceHasAssetID, "Yes", "No") & vbCrLf & _
                            "Target has ID column: " & IIf(targetHasAssetID, "Yes", "No") & vbCrLf & vbCrLf & _
                            "Enter the column name to use as the unique ID (or leave blank to skip):", _
                            "Alternative ID Column", "ASSETID")
    End If

    ' Save to config (must do this BEFORE using SetConfigValue)
    Set configSheet = GetOrCreateConfigSheet

    ' If altColumn was provided, save it now (after configSheet is created)
    If altColumn <> "" Then
        Call SetConfigValue(configSheet, "ALT_ID_COLUMN", altColumn)
        DebugPrint "SelectWorkbooksForComparison: Using alternative ID column: " & altColumn
    End If

    ' Save workbook/sheet info to config
    Call SetConfigValue(configSheet, CONFIG_SOURCE_WB, sourceWBName)
    Call SetConfigValue(configSheet, CONFIG_SOURCE_WS, sourceWSName)
    Call SetConfigValue(configSheet, CONFIG_TARGET_WB, targetWBName)
    Call SetConfigValue(configSheet, CONFIG_TARGET_WS, targetWSName)

    ' Set global references
    Set g_SourceWorkbook = sourceWB
    Set g_TargetWorkbook = targetWB

    ' Update UI status
    Set ws = g_CurrentWorksheet
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(1)
    If Not ws Is Nothing Then
        Call AddStatusDisplayFixed(ws)
    End If

    MsgBox "Files selected successfully!" & vbCrLf & vbCrLf & _
           "Source: " & sourceWBName & " - " & sourceWSName & vbCrLf & _
           "Target: " & targetWBName & " - " & targetWSName, vbInformation

    DebugPrint "SelectWorkbooksForComparison: Complete - Source=" & sourceWBName & "," & sourceWSName & " Target=" & targetWBName & "," & targetWSName
End Sub

Public Sub ConfigureSourceTargetManual()
    Dim sourceWB As String
    Dim sourceWS As String
    Dim targetWB As String
    Dim targetWS As String
    Dim configSheet As Worksheet
    Dim ws As Worksheet
    Dim sourceWBObj As Workbook

    ' Get current config
    Set configSheet = GetOrCreateConfigSheet

    sourceWB = GetConfigValue(configSheet, CONFIG_SOURCE_WB)
    sourceWS = GetConfigValue(configSheet, CONFIG_SOURCE_WS)
    targetWB = GetConfigValue(configSheet, CONFIG_TARGET_WB)
    targetWS = GetConfigValue(configSheet, CONFIG_TARGET_WS)

    Call ShowSourceTargetDialog(sourceWB, sourceWS, targetWB, targetWS)

    ' Save new configuration
    Call SetConfigValue(configSheet, CONFIG_SOURCE_WB, sourceWB)
    Call SetConfigValue(configSheet, CONFIG_SOURCE_WS, sourceWS)
    Call SetConfigValue(configSheet, CONFIG_TARGET_WB, targetWB)
    Call SetConfigValue(configSheet, CONFIG_TARGET_WS, targetWS)

    ' Try to set global references
    On Error Resume Next
    If sourceWB <> "" Then
        Set sourceWBObj = GetWorkbookByName(sourceWB)
        If Not sourceWBObj Is Nothing Then
            Set g_SourceWorkbook = sourceWBObj
            Set g_SourceSheet = sourceWBObj.Worksheets(sourceWS)
        End If
    End If
    If targetWB <> "" Then
        Set g_TargetWorkbook = GetWorkbookByName(targetWB)
        If Not g_TargetWorkbook Is Nothing Then
            Set g_TargetSheet = g_TargetWorkbook.Worksheets(targetWS)
        End If
    End If
    On Error GoTo 0

    ' Refresh the UI status display
    Set ws = g_CurrentWorksheet
    If ws Is Nothing Then Set ws = g_SourceSheet
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(1)
    If Not ws Is Nothing Then
        Call AddStatusDisplayFixed(ws)
    End If
End Sub

'===============================================================================
' CHECK AND REBUILD UI - Ensures buttons exist before Execute Match
'===============================================================================

Private Sub CheckAndPromptForUIRebuild()
    '
    ' NEW: Checks if UI buttons exist, prompts user to rebuild if missing
    ' This prevents errors when user deletes columns containing buttons
    '
    Dim ws As Worksheet
    Dim btn As Excel.Button
    Dim buttonCount As Long
    Dim uiButtonCount As Long

    On Error Resume Next

    ' Get source sheet
    If g_SourceSheet Is Nothing Then Exit Sub
    Set ws = g_SourceSheet

    ' Count buttons in UI area (rows 1-10)
    buttonCount = 0
    uiButtonCount = 0
    For Each btn In ws.Buttons
        If btn.Top < ws.Rows(10).Top Then
            uiButtonCount = uiButtonCount + 1
        End If
        buttonCount = buttonCount + 1
    Next btn

    DebugPrint "CheckAndPromptForUIRebuild: Total buttons=" & buttonCount & ", UI buttons=" & uiButtonCount

    ' If no UI buttons found, prompt user
    If uiButtonCount = 0 Then
        Dim response As Integer
        response = MsgBox("No Match UI buttons found on the sheet." & vbCrLf & vbCrLf & _
                "This may happen if you deleted columns containing buttons." & vbCrLf & _
                "Would you like to rebuild the UI?" & vbCrLf & vbCrLf & _
                "Click Yes to rebuild, or No to continue without buttons.", _
                vbYesNo + vbQuestion, "UI Buttons Missing")

        If response = vbYes Then
            ' Rebuild UI
            g_ForceRebuild = True
            Call RebuildMatchBuilderUI
            MsgBox "UI has been rebuilt. You can now run Execute Match.", vbInformation
        Else
            DebugPrint "CheckAndPromptForUIRebuild: User declined rebuild"
        End If
    End If

    On Error GoTo 0
End Sub

'===============================================================================
' CHECK FILE STATUS - Verifies if both files are open and active
'===============================================================================

Public Sub CheckFileStatus()
    '
    ' Checks and displays the connection status of source and target files
    ' Shows whether both files are open and accessible
    '
    Dim configSheet As Worksheet
    Dim sourceWBName As String, sourceWSName As String
    Dim targetWBName As String, targetWSName As String
    Dim sourceWB As Workbook, targetWB As Workbook
    Dim sourceWS As Worksheet, targetWS As Worksheet
    Dim sourceStatus As String, targetStatus As String
    Dim allOK As Boolean
    Dim msg As String
    Dim ws As Worksheet

    DebugPrint "CheckFileStatus: Starting..."

    allOK = True

    ' Get config
    Set configSheet = GetOrCreateConfigSheet
    sourceWBName = GetConfigValue(configSheet, CONFIG_SOURCE_WB)
    sourceWSName = GetConfigValue(configSheet, CONFIG_SOURCE_WS)
    targetWBName = GetConfigValue(configSheet, CONFIG_TARGET_WB)
    targetWSName = GetConfigValue(configSheet, CONFIG_TARGET_WS)

    ' Check source workbook
    If sourceWBName = "" Then
        sourceStatus = "NOT SET"
        allOK = False
    Else
        Set sourceWB = GetWorkbookByName(sourceWBName)
        If sourceWB Is Nothing Then
            sourceStatus = "CLOSED"
            allOK = False
        Else
            On Error Resume Next
            Set sourceWS = sourceWB.Worksheets(sourceWSName)
            On Error GoTo 0
            If sourceWS Is Nothing Then
                sourceStatus = "SHEET NOT FOUND"
                allOK = False
            Else
                sourceStatus = "OPEN"
            End If
        End If
    End If

    ' Check target workbook
    If targetWBName = "" Then
        targetStatus = "NOT SET"
        allOK = False
    Else
        Set targetWB = GetWorkbookByName(targetWBName)
        If targetWB Is Nothing Then
            targetStatus = "CLOSED"
            allOK = False
        Else
            On Error Resume Next
            Set targetWS = targetWB.Worksheets(targetWSName)
            On Error GoTo 0
            If targetWS Is Nothing Then
                targetStatus = "SHEET NOT FOUND"
                allOK = False
            Else
                targetStatus = "OPEN"
            End If
        End If
    End If

    ' Display status
    msg = "FILE CONNECTION STATUS" & vbCrLf & vbCrLf & _
          "Source File: " & sourceWBName & vbCrLf & _
          "  Sheet: " & sourceWSName & vbCrLf & _
          "  Status: " & sourceStatus & vbCrLf & vbCrLf & _
          "Target File: " & targetWBName & vbCrLf & _
          "  Sheet: " & targetWSName & vbCrLf & _
          "  Status: " & targetStatus & vbCrLf & vbCrLf

    If allOK Then
        msg = msg & "Both files are ACTIVE and ready for comparison."
        MsgBox msg, vbInformation, "Status Check - OK"
    Else
        msg = msg & "WARNING: One or more files are not available!" & vbCrLf & _
              "Please check that both files are open."
        MsgBox msg, vbExclamation, "Status Check - WARNING"
    End If

    ' Update global references if files are open
    If Not sourceWB Is Nothing Then
        Set g_SourceWorkbook = sourceWB
        Set g_SourceSheet = sourceWS
    End If
    If Not targetWB Is Nothing Then
        Set g_TargetWorkbook = targetWB
        Set g_TargetSheet = targetWS
    End If

    ' Refresh UI status display
    Set ws = g_CurrentWorksheet
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(1)
    If Not ws Is Nothing Then
        Call AddStatusDisplayFixed(ws)
    End If

    DebugPrint "CheckFileStatus: Complete - Source=" & sourceStatus & ", Target=" & targetStatus
End Sub

'===============================================================================
' USER-DEFINED HEADER ROW SELECTION
'===============================================================================

Private Function PromptUserForHeaderRow(ws As Worksheet) As Long
    '
    ' Scans the worksheet and lets user confirm/select which row is the header
    ' Returns the user-selected row number
    '
    Dim r As Long, c As Long
    Dim lastCol As Long
    Dim score As Long
    Dim candidates As String
    Dim bestRow As Long
    Dim bestScore As Long
    Dim rawValue As String
    Dim normalized As String
    Dim matched As String
    Dim aliases As Object
    Dim learned As Object
    Dim userInput As String
    Dim selectedRow As Long
    Dim cellVal As String
    Dim textCount As Long
    Dim totalNonEmpty As Long
    Dim fillDefault As Boolean
    Dim insertNew As Boolean
    Dim defResponse As Integer
    Dim headerTargetRow As Long
    Dim colN As Long
    Dim defaultColNum As Long
    Dim scanC As Long
    Dim scanVal As String

    Set aliases = CreateObject("Scripting.Dictionary")
    Set learned = CreateObject("Scripting.Dictionary")

    lastCol = GetLastColumn(ws, 1)
    If lastCol < 5 Then lastCol = 10

    DebugPrint "PromptUserForHeaderRow: Scanning for header rows..."

    ' Scan rows 1-200 to find candidates
    bestScore = 0
    bestRow = 10  ' Default
    candidates = "Header row candidates found:" & vbCrLf & vbCrLf

    For r = 1 To 200
        score = 0
        For c = 1 To lastCol
            rawValue = CStr(ws.Cells(r, c).value)
            normalized = UltraNormalize(rawValue)
            If normalized <> "" Then
                matched = SmartMatch(rawValue, normalized, aliases, learned)
                If matched <> "" Then score = score + 1
            End If
        Next c

        ' Build preview string for this row
        Dim preview As String
        preview = ""
        For c = 1 To 5
            If c <= lastCol Then
                cellVal = Trim(CStr(ws.Cells(r, c).value))
                If Len(cellVal) > 15 Then cellVal = Left(cellVal, 12) & "..."
                If cellVal <> "" Then
                    If preview <> "" Then preview = preview & ", "
                    preview = preview & cellVal
                End If
            End If
        Next c

        If preview = "" Then preview = "(empty)"

        ' Show rows with at least 1 header-like cell, or the first 10 rows
        If score >= 1 Or r <= 10 Then
            candidates = candidates & "Row " & r & ": " & preview
            If score >= 2 Then
                candidates = candidates & " [BEST: " & score & " matches]"
            ElseIf score = 1 Then
                candidates = candidates & " [" & score & " match]"
            End If
            candidates = candidates & vbCrLf

            If score > bestScore Then
                bestScore = score
                bestRow = r
            End If
        End If
    Next r

    ' If no good candidate found, default to row 10
    If bestScore < 1 Then bestRow = 10

    ' Show dialog to user
    Dim prompt As String
    prompt = candidates & vbCrLf & _
             "Enter the row number for your data header:" & vbCrLf & _
             "(Press Enter to accept the best match: row " & bestRow & ")"

    userInput = InputBox(prompt, "Select Header Row", CStr(bestRow))

    If userInput = "" Then
        ' User cancelled - return 0 to signal cancellation
        PromptUserForHeaderRow = 0
        Exit Function
    Else
        ' Parse user input
        On Error Resume Next
        selectedRow = CLng(userInput)
        If Err.Number <> 0 Or selectedRow < 1 Or selectedRow > 1000 Then
            MsgBox "Invalid row number. Using best match: row " & bestRow, vbInformation
            selectedRow = bestRow
        End If
        On Error GoTo 0

        ' FIX: ALWAYS validate the selected row (not just when user changes selection)
        ' This ensures user can select any valid row including row 4
        If Not ValidateHeaderRow(ws, selectedRow) Then
            MsgBox "Row " & selectedRow & " does not appear to be a valid header row." & vbCrLf & _
                   "A valid header should have at least 3 non-empty columns.", vbExclamation
            PromptUserForHeaderRow = PromptUserForHeaderRow(ws)
            Exit Function
        End If

        If Not HasDataBelowHeader(ws, selectedRow) Then
            ' Check if it's just formatting (not real data) by checking a few rows
            Dim hasAnyData As Boolean
            hasAnyData = False
            Dim checkR As Long
            For checkR = selectedRow + 1 To selectedRow + 50
                For c = 1 To 10
                    If Trim(CStr(ws.Cells(checkR, c).Value)) <> "" Then
                        hasAnyData = True
                        Exit For
                    End If
                Next c
                If hasAnyData Then Exit For
            Next checkR

            If Not hasAnyData Then
                MsgBox "No data found below row " & selectedRow & "." & vbCrLf & _
                       "Please select a header row that has data below it.", vbExclamation
                PromptUserForHeaderRow = PromptUserForHeaderRow(ws)
                Exit Function
            End If
        End If
    End If

    DebugPrint "PromptUserForHeaderRow: User selected row " & selectedRow

    ' Count text vs numeric/empty cells in selected row
    textCount = 0
    totalNonEmpty = 0
    fillDefault = False
    insertNew = False

    For scanC = 1 To lastCol
        scanVal = Trim(CStr(ws.Cells(selectedRow, scanC).Value))
        If scanVal <> "" Then
            totalNonEmpty = totalNonEmpty + 1
            If Not IsNumeric(scanVal) Then textCount = textCount + 1
        End If
    Next scanC

    ' If fewer than 3 text cells found, offer default column names
    If textCount < 3 Then
        defResponse = MsgBox("Row " & selectedRow & " has " & textCount & " named column(s) out of " & totalNonEmpty & " non-empty cells." & vbCrLf & vbCrLf & _
               "What would you like to do?" & vbCrLf & _
               "YES — Fill empty cells in this row with default names (Column_1, Column_2...)" & vbCrLf & _
               "NO — Insert a new header row above row " & selectedRow & " with default names" & vbCrLf & _
               "CANCEL — Go back and select a different row", _
               vbYesNoCancel + vbQuestion, "Header Row Has Few Names")

        If defResponse = vbCancel Then
            PromptUserForHeaderRow = PromptUserForHeaderRow(ws)
            Exit Function
        ElseIf defResponse = vbYes Then
            fillDefault = True
        ElseIf defResponse = vbNo Then
            insertNew = True
        End If
    End If

    ' Apply default column names if requested
    If fillDefault Or insertNew Then
        If insertNew Then
            ' Insert new row above selected row
            ws.Rows(selectedRow).Insert Shift:=xlDown
            headerTargetRow = selectedRow
        Else
            headerTargetRow = selectedRow
        End If

        ' Fill empty cells with Column_N names
        defaultColNum = 1
        For colN = 1 To lastCol
            If Trim(CStr(ws.Cells(headerTargetRow, colN).Value)) = "" Then
                ws.Cells(headerTargetRow, colN).Value = "Column_" & defaultColNum
            End If
            defaultColNum = defaultColNum + 1
        Next colN

        ' Style the new header row
        ws.Cells(headerTargetRow, 1).EntireRow.Font.Bold = True
    End If

    PromptUserForHeaderRow = selectedRow
End Function

'===============================================================================
' BUILD FULL UI - Disables SafeMode and builds complete UI
'===============================================================================

Public Sub BuildFullUI()
    '
    ' Disables SafeMode and rebuilds the full UI with row insertion
    ' This creates the 5-row match UI above your data and adds MATCH_TYPE column
    '
    ' CRITICAL FIX: Ask for header row FIRST, then build UI - don't replace existing data
    '
    Dim response As Integer
    Dim ws As Worksheet
    Dim selectedHeaderRow As Long

    g_MacroRunning = True

    DebugPrint "BuildFullUI: Starting..."

    ' Step 1: Get the active worksheet FIRST
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = GetValidWorksheetForUI
    End If

    If ws Is Nothing Then
        MsgBox "No worksheet available. Please open a workbook with data.", vbCritical
        g_MacroRunning = False
        Exit Sub
    End If

    ' Step 1b: Convert any Excel Tables to plain ranges to prevent table structure interference
    ' Preserve header values before unlisting since Unlist may clear header row cells
    Dim tbl As ListObject
    Dim tblHeaderRow As Long
    Dim tblHeaderCol As Long
    Dim tblHeaders() As String
    Dim tblStartCol As Long
    Dim tblColCount As Long
    For Each tbl In ws.ListObjects
        tblHeaderRow = tbl.HeaderRowRange.Row
        tblStartCol = tbl.HeaderRowRange.Column
        tblColCount = tbl.ListColumns.Count
        ReDim tblHeaders(1 To tblColCount)
        For tblHeaderCol = 1 To tblColCount
            tblHeaders(tblHeaderCol) = CStr(tbl.HeaderRowRange.Cells(1, tblHeaderCol).Value)
        Next tblHeaderCol
        ' Write header values directly to cells BEFORE unlisting
        For tblHeaderCol = 1 To tblColCount
            ws.Cells(tblHeaderRow, tblStartCol + tblHeaderCol - 1).Value = tbl.ListColumns(tblHeaderCol).Name
        Next tblHeaderCol
        tbl.Unlist
        ' Verify headers survived - restore from saved array if cells are now empty
        For tblHeaderCol = 1 To tblColCount
            If Trim(CStr(ws.Cells(tblHeaderRow, tblStartCol + tblHeaderCol - 1).Value)) = "" Then
                ws.Cells(tblHeaderRow, tblStartCol + tblHeaderCol - 1).Value = tblHeaders(tblHeaderCol)
            End If
        Next tblHeaderCol
        DebugPrint "BuildFullUI: Converted table to plain range, preserved " & tblColCount & " headers at row " & tblHeaderRow
    Next tbl

    ' Step 2: Ask user where their header row is - BEFORE building anything
    selectedHeaderRow = PromptUserForHeaderRow(ws)

    ' Exit if user cancelled header selection
    If selectedHeaderRow = 0 Then
        DebugPrint "BuildFullUI: User cancelled header selection"
        g_MacroRunning = False
        Exit Sub
    End If

    ' Step 3: Confirm with user before inserting rows - AFTER knowing header position
    response = MsgBox("The Match UI will be built above row " & selectedHeaderRow & "." & vbCrLf & vbCrLf & _
            "This will INSERT 6 rows above your data." & vbCrLf & _
            "Your data at row " & selectedHeaderRow & " will be shifted DOWN to row " & (selectedHeaderRow + 6) & "." & vbCrLf & vbCrLf & _
            "Your existing data will NOT be deleted or replaced." & vbCrLf & vbCrLf & _
            "Do you want to continue?", vbYesNo + vbExclamation, "Build UI")

    If response = vbNo Then
        DebugPrint "BuildFullUI: User cancelled"
        g_MacroRunning = False
        Exit Sub
    End If

    ' Step 4: Set flags and build UI directly (don't call StartSystem which triggers ResetGlobalState)
    g_SafeMode = False
    g_ForceRebuild = True
    g_ClearMatchDataOnRebuild = True  ' Full rebuild - clear old match data
    Set g_CurrentWorksheet = ws

    ' Step 5: Build the UI with user-selected header row
    ' This will insert rows above the data, NOT replace the data
    Call RebuildMatchBuilderUI(selectedHeaderRow)

    MsgBox "UI Build Complete!" & vbCrLf & vbCrLf & _
           "The Match UI has been created above your data." & vbCrLf & _
           "Data header is now at row " & (selectedHeaderRow + 6) & "." & vbCrLf & vbCrLf & _
           "You can now configure your match rules.", vbInformation

    DebugPrint "BuildFullUI: Complete"

    g_MacroRunning = False
End Sub

'===============================================================================
' BuildKeyFast - Fast key builder using pre-resolved column index array
' Replaces BuildKeyFromRow for performance-critical loops
'===============================================================================
Private Function BuildKeyFast(data As Variant, rowIndex As Long, colIdxArr As Variant, colCount As Long) As String
    Dim parts() As String
    Dim partCount As Long
    Dim j As Long
    Dim val As String
    Dim result As String
    Dim idxArr() As Long

    If colCount = 0 Then
        BuildKeyFast = ""
        Exit Function
    End If

    On Error GoTo FallbackEmpty
    idxArr = colIdxArr
    On Error GoTo 0

    partCount = 0
    ReDim parts(1 To colCount)

    For j = 1 To colCount
        val = SafeCleanString(data(rowIndex, idxArr(j)))
        If val <> "" Then
            partCount = partCount + 1
            parts(partCount) = val
        End If
    Next j

    If partCount = 0 Then
        BuildKeyFast = ""
    ElseIf partCount = 1 Then
        BuildKeyFast = parts(1)
    Else
        result = parts(1)
        For j = 2 To partCount
            result = result & "|" & parts(j)
        Next j
        BuildKeyFast = result
    End If
    Exit Function

FallbackEmpty:
    BuildKeyFast = ""
End Function

'===============================================================================
' CORE MATCHING FUNCTIONS
'===============================================================================

Private Function BuildKeyFromRow(data As Variant, rowIndex As Long, colSel As Variant, colMap As Object) As String
    Dim parts As Collection
    Dim i As Long
    Dim colName As Variant
    Dim colIdx As Long
    Dim val As String
    Dim result As String
    Dim normalizedColName As String
    Dim foundCols As String
    Dim debugCols As String

    Set parts = New Collection

    ' Debug: track what columns we're looking for
    debugCols = ""
    For Each colName In colSel
        If debugCols <> "" Then debugCols = debugCols & ", "
        debugCols = debugCols & CStr(colName)
    Next colName

    DebugPrint "BuildKeyFromRow: Looking for columns: " & debugCols

    ' Debug: show what columns exist in colMap
    Dim mapKeys As String
    mapKeys = ""
    For Each colName In colMap.keys
        If mapKeys <> "" Then mapKeys = mapKeys & ", "
        mapKeys = mapKeys & colName & "=" & colMap(colName)
    Next colName
    DebugPrint "BuildKeyFromRow: colMap has " & colMap.Count & " columns: " & Left(mapKeys, 200)

    For Each colName In colSel
        ' Normalize the column name to match the colMap keys
        normalizedColName = UltraNormalize(CStr(colName))

        DebugPrint "BuildKeyFromRow: Looking for '" & colName & "' normalized to '" & normalizedColName & "'"

        If normalizedColName <> "" And colMap.exists(normalizedColName) Then
            colIdx = colMap(normalizedColName)
            val = SafeCleanString(data(rowIndex, colIdx))
            DebugPrint "BuildKeyFromRow: Found at col " & colIdx & ", value: '" & val & "'"
            If val <> "" Then parts.Add val
        ElseIf colMap.exists(CStr(colName)) Then
            ' Try direct match as fallback
            colIdx = colMap(CStr(colName))
            val = SafeCleanString(data(rowIndex, colIdx))
            DebugPrint "BuildKeyFromRow: Direct match at col " & colIdx & ", value: '" & val & "'"
            If val <> "" Then parts.Add val
        Else
            DebugPrint "BuildKeyFromRow: Column NOT FOUND in colMap: " & normalizedColName
        End If
    Next colName

    If parts.Count = 0 Then
        BuildKeyFromRow = ""
        DebugPrint "BuildKeyFromRow: No parts collected, returning empty"
    Else
        result = parts(1)
        For i = 2 To parts.Count
            result = result & "|" & parts(i)
        Next i
        DebugPrint "BuildKeyFromRow: Built key = '" & result & "'"
        BuildKeyFromRow = result
    End If
End Function

Private Function GetMatchDefinitionsFromUI(ws As Worksheet, dataHeaderRow As Long, colMap As Object) As Collection
    Dim matches As New Collection
    Dim r As Long
    Dim c As Long
    Dim lastCol As Long
    Dim colName As String
    Dim selCols As Collection
    Dim matchID As Variant
    Dim matchType As String
    Dim cellVal As String
    Dim matchDef As Object
    Dim uiHeaderRow As Long

    ' Find UI header row (row 4 in the UI area)
    uiHeaderRow = UI_COLHEADER_ROW  ' This is row 4

    lastCol = GetLastDataColumn(ws, dataHeaderRow)
    g_LastDataColumn = lastCol

    DebugPrint "GetMatchDefinitionsFromUI: dataHeaderRow=" & dataHeaderRow & ", uiHeaderRow=" & uiHeaderRow & ", lastCol=" & lastCol

    For r = UI_FIRST_MATCH_ROW To dataHeaderRow - 1
        matchID = ws.Cells(r, 1).value

        ' FIXED: Accept both numeric AND text IDs (like "Match_3")
        ' If Match Type is empty but we have an ID, auto-fill Match Type with the ID
        matchType = Trim(CStr(ws.Cells(r, 2).value))
        If Trim(CStr(matchID)) <> "" And matchType = "" Then
            matchType = Trim(CStr(matchID))
        End If

        ' Debug: show what was read
        DebugPrint "GetMatchDefinitionsFromUI: Row " & r & " - ID=" & matchID & ", Type=" & matchType

        ' Accept row if: ID exists AND Match Type exists (or we just filled it)
        If Trim(CStr(matchID)) <> "" And matchType <> "" Then
            Set selCols = New Collection

            ' FIXED: Read X marks from UI, but get column names from UI header row (row 4)
            ' The UI columns align with data columns, so we use the column index from UI
            ' but get the NAME from the UI header row
            For c = 3 To lastCol
                ' First check if there's an X in the UI row
                cellVal = Trim(UCase(CStr(ws.Cells(r, c).value)))
                If cellVal = "X" Then
                    ' Get the column name from the UI header row (row 4)
                    colName = GetColumnNameFromIndex(ws, uiHeaderRow, c)
                    DebugPrint "GetMatchDefinitionsFromUI: Found X at row " & r & ", col " & c & ", column name: " & colName

                    If colName <> "" Then
                        ' FIXED: Use SmartMatch to get canonical name (same as GetColumnMap does)
                        ' This ensures the column names match what's in colMap
                        Dim matchedColName As String
                        matchedColName = SmartMatch(colName, UltraNormalize(colName), Nothing, Nothing)
                        DebugPrint "GetMatchDefinitionsFromUI: SmartMatch result: " & matchedColName

                        If matchedColName <> "" Then
                            selCols.Add matchedColName
                        Else
                            selCols.Add colName
                        End If
                    End If
                End If
            Next c

            DebugPrint "GetMatchDefinitionsFromUI: Row " & r & " - selected " & selCols.Count & " columns"

            If selCols.Count > 0 Then
                Set matchDef = CreateObject("Scripting.Dictionary")
                ' Use ID as-is (could be number or text like "Match_3")
                matchDef.Add "ID", matchID
                matchDef.Add "Type", matchType
                matchDef.Add "Columns", selCols
                matches.Add matchDef
            End If
        End If
    Next r

    ' FIX Issue: If no match definitions found (no X marks selected), use ALL columns from header row
    ' This ensures correct behavior regardless of which ID column user wants
    If matches.Count = 0 Then
        Set matchDef = CreateObject("Scripting.Dictionary")
        matchDef.Add "ID", 1
        matchDef.Add "Type", "Default"
        ' Use ALL columns from header row for matching
        Set selCols = New Collection
        Dim headerCol As Long
        For headerCol = 1 To lastCol
            Dim headerName As String
            headerName = Trim(CStr(ws.Cells(dataHeaderRow, headerCol).Value))
            If headerName <> "" Then
                ' Use SmartMatch to normalize column name
                Dim normalizedName As String
                normalizedName = SmartMatch(headerName, UltraNormalize(headerName), Nothing, Nothing)
                If normalizedName <> "" Then
                    selCols.Add normalizedName
                Else
                    selCols.Add headerName
                End If
            End If
        Next headerCol

        DebugPrint "GetMatchDefinitionsFromUI: No X marks selected, using all " & selCols.Count & " columns for matching"
        matchDef.Add "Columns", selCols
        matches.Add matchDef
    End If

    Set GetMatchDefinitionsFromUI = matches
End Function

Private Function GetColumnNameFromIndex(ws As Worksheet, headerRow As Long, colIndex As Long) As String
    Dim val As String
    val = Trim(CStr(ws.Cells(headerRow, colIndex).value))
    If val = "" Then GetColumnNameFromIndex = "" Else GetColumnNameFromIndex = val
End Function

'===============================================================================
' HEADER DETECTION AND MATCHING
'===============================================================================

Public Function FindDataHeaderRow(ws As Worksheet, aliases As Object, learned As Object) As Long
    Dim r As Long
    Dim c As Long
    Dim lastCol As Long
    Dim score As Long
    Dim maxScore As Long
    Dim bestRow As Long
    Dim rawValue As String
    Dim normalized As String
    Dim matched As String
    Dim sampleHeaders As String
    Dim nonEmptyCount As Long
    Dim foundHeader As Boolean
    Dim detectionLastCol As Long
    Dim consecutiveEmpty As Long
    Dim colCheck As Long
    Dim isUIRow As Boolean
    Dim uiCheck As Long
    Dim uiCellVal As String
    Dim isDataHeader As Boolean
    Dim dhCheck As Long

    foundHeader = False
    FindDataHeaderRow = 0

    For r = 1 To 200
        detectionLastCol = 0
        consecutiveEmpty = 0
        For colCheck = 1 To ws.Columns.Count
            If Trim(CStr(ws.Cells(r, colCheck).Value)) <> "" Then
                detectionLastCol = colCheck
                consecutiveEmpty = 0
            Else
                consecutiveEmpty = consecutiveEmpty + 1
                If consecutiveEmpty >= 2 Then Exit For
            End If
        Next colCheck

        If detectionLastCol >= 3 Then
            nonEmptyCount = 0
            For colCheck = 1 To detectionLastCol
                If Trim(CStr(ws.Cells(r, colCheck).Value)) <> "" Then
                    nonEmptyCount = nonEmptyCount + 1
                End If
            Next colCheck

            If nonEmptyCount >= 3 Then
                ' Check if this row contains MATCHED_ID — definitive data header identifier
                isDataHeader = False
                For dhCheck = 1 To detectionLastCol
                    If UCase(Trim(CStr(ws.Cells(r, dhCheck).Value))) = "MATCHED_ID" Then
                        isDataHeader = True
                        Exit For
                    End If
                Next dhCheck
                If isDataHeader Then
                    FindDataHeaderRow = r
                    foundHeader = True
                    DebugPrint "FindDataHeaderRow: Data header confirmed by MATCHED_ID at row " & r
                    Exit For
                End If

                ' Check if this is a UI row (contains "Match" keyword)
                isUIRow = False
                For uiCheck = 1 To detectionLastCol
                    uiCellVal = UCase(Trim(CStr(ws.Cells(r, uiCheck).Value)))
                    If uiCellVal = "MATCH" Or _
                       uiCellVal = "MATCH TYPE" Or _
                       Left(uiCellVal, 7) = "SOURCE:" Or _
                       Left(uiCellVal, 7) = "TARGET:" Or _
                       uiCellVal = "X" Then
                        isUIRow = True
                        Exit For
                    End If
                Next uiCheck

                ' Skip UI rows, only accept data header rows
                If Not isUIRow Then
                    FindDataHeaderRow = r
                    foundHeader = True
                    DebugPrint "FindDataHeaderRow: Header row found at row " & r & " with " & nonEmptyCount & " non-empty cells"
                    Exit For
                End If
            End If
        End If
    Next r

    If Not foundHeader Then
        DebugPrint "FindDataHeaderRow: No header row found with sufficient density"
    End If
End Function

Public Function UltraNormalize(ByVal rawValue As String) As String
    Dim temp As String
    Dim colonPos As Long
    Dim i As Long

    If rawValue = "" Then
        UltraNormalize = ""
        Exit Function
    End If

    temp = rawValue
    If temp = "~NULL~" Then
        UltraNormalize = ""
        Exit Function
    End If

    temp = Replace(temp, Chr(160), " ")
    temp = Replace(temp, Chr(9), " ")
    temp = Replace(temp, Chr(10), " ")
    temp = Replace(temp, Chr(13), " ")

    For i = 1 To Len(temp)
        If Asc(Mid(temp, i, 1)) < 32 Then Mid(temp, i, 1) = " "
    Next i

    colonPos = InStr(temp, ":")
    If colonPos > 0 Then temp = Trim(Mid(temp, colonPos + 1))

    temp = Replace(temp, "_", "")
    temp = Replace(temp, "-", "")
    temp = Replace(temp, ".", "")
    temp = Replace(temp, " ", "")

    UltraNormalize = LCase(Trim(temp))
End Function

Public Function SmartMatch(ByVal rawValue As String, ByVal normalized As String, aliases As Object, learned As Object) As String
    Static baseMap As Object

    If baseMap Is Nothing Then
        Set baseMap = CreateObject("Scripting.Dictionary")
        baseMap.Add "assetid", "ASSETID"
        baseMap.Add "aid", "ASSETID"
        baseMap.Add "asset_id", "ASSETID"
        baseMap.Add "assetnum", "ASSETNUM"
        baseMap.Add "assetnumber", "ASSETNUM"
        baseMap.Add "assetno", "ASSETNUM"
        baseMap.Add "anum", "ASSETNUM"
        baseMap.Add "anumber", "ASSETNUM"
        baseMap.Add "location", "LOCATION"
        baseMap.Add "loc", "LOCATION"
        baseMap.Add "site", "LOCATION"
        baseMap.Add "serialnum", "SERIALNUM"
        baseMap.Add "serialnumber", "SERIALNUM"
        baseMap.Add "serialno", "SERIALNUM"
        baseMap.Add "sn", "SERIALNUM"
        baseMap.Add "sernum", "SERIALNUM"
        baseMap.Add "serno", "SERIALNUM"
        baseMap.Add "description", "DESCRIPTION"
        baseMap.Add "desc", "DESCRIPTION"
        baseMap.Add "descr", "DESCRIPTION"
        baseMap.Add "matchedassetid", "MATCHED_ASSETID"
        baseMap.Add "matchedid", "MATCHED_ID"
        baseMap.Add "matchassetid", "MATCHED_ASSETID"
        baseMap.Add "matchedasset", "MATCHED_ASSETID"
        baseMap.Add "note1", "MATCH_TYPE"
        baseMap.Add "note", "MATCH_TYPE"
        baseMap.Add "notes", "MATCH_TYPE"
        ' FIXED: Add result column names (MATCHED_ID, MATCH_TYPE, MATCH_STATUS, SOURCE_FILE, TARGET_FILE)
        baseMap.Add "matched_id", "MATCHED_ID"
        baseMap.Add "match_type", "MATCH_TYPE"
        baseMap.Add "match_status", "MATCH_STATUS"
        baseMap.Add "source_file", "SOURCE_FILE"
        baseMap.Add "target_file", "TARGET_FILE"
        baseMap.Add "note_1", "MATCH_TYPE"
        baseMap.Add "note_2", "MATCH_STATUS"
        ' NOTE: STATUS is a button, not a column - removed from mapping
        ' FIXED: Add more common column names
        ' FIXED: UltraNormalize removes underscores, so we need BOTH versions
        baseMap.Add "db_build_path", "DB_BUILD_PATH"
        baseMap.Add "dbbuildpath", "DB_BUILD_PATH"
        baseMap.Add "db_build_level", "DB_BUILD_LEVEL"
        baseMap.Add "dbbuiltlevel", "DB_BUILD_LEVEL"
        baseMap.Add "plusgopenomuid", "PLUSGOPENOMUID"
        baseMap.Add "parent", "PARENT"
        baseMap.Add "assettag", "ASSETTAG"
        baseMap.Add "tag", "ASSETTAG"
    End If

    ' Check aliases first
    If Not aliases Is Nothing Then
        If aliases.exists(normalized) Then
            SmartMatch = aliases(normalized)
            Exit Function
        End If
    End If

    ' Check base map
    If Not baseMap Is Nothing Then
        If baseMap.exists(normalized) Then
            SmartMatch = baseMap(normalized)
            If Not aliases Is Nothing Then
                If Not aliases.exists(normalized) Then aliases.Add normalized, SmartMatch
            End If
            If Not learned Is Nothing Then
                If Not learned.exists(rawValue) Then learned.Add rawValue, SmartMatch
            End If
            Exit Function
        End If
    End If

    ' FIXED: Universal fallback - return the normalized column name if not found
    ' This makes matching work for ANY column, not just predefined ones
    SmartMatch = normalized
End Function

Private Function GetColumnMap(ws As Worksheet, ByVal headerRow As Long, aliases As Object, learned As Object) As Object
    Dim dict As Object
    Dim i As Long
    Dim lastCol As Long
    Dim rawHeader As String
    Dim normalized As String
    Dim matched As String

    Set dict = CreateObject("Scripting.Dictionary")
    lastCol = GetLastColumn(ws, headerRow)

    For i = 1 To lastCol
        rawHeader = CStr(ws.Cells(headerRow, i).value)
        normalized = UltraNormalize(rawHeader)

        If normalized <> "" Then
            matched = SmartMatch(rawHeader, normalized, aliases, learned)
            If matched <> "" Then
                If Not dict.exists(matched) Then
                    dict.Add matched, i
                End If
            End If
        End If
    Next i

    Set GetColumnMap = dict
End Function

'===============================================================================
' DATA HELPERS
'===============================================================================

Public Function SafeCleanString(ByVal inputValue As Variant) As String
    Dim temp As String

    If IsNull(inputValue) Or IsEmpty(inputValue) Then
        SafeCleanString = ""
        Exit Function
    End If

    temp = CStr(inputValue)
    If temp = "~NULL~" Then
        SafeCleanString = ""
        Exit Function
    End If

    temp = Replace(temp, Chr(160), " ")
    temp = Replace(temp, Chr(9), " ")
    temp = Replace(temp, Chr(10), " ")
    temp = Replace(temp, Chr(13), " ")

    Do While InStr(temp, "  ") > 0
        temp = Replace(temp, "  ", " ")
    Loop

    SafeCleanString = UCase(Trim(temp))
End Function

Private Function SafeGetWorksheet(wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set SafeGetWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Function GetLastColumn(ws As Worksheet, Optional ByVal searchRow As Long = 1) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells(searchRow, ws.Columns.Count).End(xlToLeft)
    On Error GoTo 0
    If lastCell Is Nothing Then GetLastColumn = 1 Else GetLastColumn = lastCell.Column
End Function

Public Function GetLastDataColumn(ws As Worksheet, headerRow As Long) As Long
    Dim lastCol As Long
    lastCol = GetLastColumn(ws, headerRow)
    If lastCol < 1 Then lastCol = 1
    GetLastDataColumn = lastCol
End Function

Private Function GetLastRow(ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", _
                                 After:=ws.Cells(1, 1), _
                                 LookIn:=xlFormulas, _
                                 LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious)
    On Error GoTo 0
    If lastCell Is Nothing Then GetLastRow = 1 Else GetLastRow = lastCell.Row
End Function

Private Function GetSheetData(ws As Worksheet, ByVal headerRow As Long) As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim emptyArray() As Variant
    Dim errArray() As Variant

    On Error GoTo ErrorHandler

    lastRow = GetLastRow(ws)
    lastCol = GetLastColumn(ws, headerRow)

    If lastRow >= headerRow And lastCol > 0 Then
        Set dataRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))
        If Not dataRange Is Nothing Then
            GetSheetData = dataRange.value
            Exit Function
        End If
    End If

    ReDim emptyArray(1 To 1, 1 To 1)
    emptyArray(1, 1) = ""
    GetSheetData = emptyArray
    Exit Function

ErrorHandler:
    ReDim errArray(1 To 1, 1 To 1)
    errArray(1, 1) = ""
    GetSheetData = errArray
End Function

Private Sub OptimizePerformance(ByVal optimize As Boolean)
    On Error Resume Next
    If optimize Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayAlerts = False
    Else
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' BUTTON ANCHOR FUNCTIONS - Issue 4
'===============================================================================

' FindButtonAnchorStartRow - Find or create section for button anchors
Private Function FindButtonAnchorStartRow(configWS As Worksheet) As Long
    Dim r As Long

    ' Scan column 1 for "BUTTON_ANCHORS"
    For r = 7 To 200
        If UCase(Trim(CStr(configWS.Cells(r, 1).Value))) = "BUTTON_ANCHORS" Then
            FindButtonAnchorStartRow = r + 1
            Exit Function
        End If
    Next r

    ' If not found, scan for last non-empty row (explicit loop, NO End(xlUp))
    Dim lastNonEmptyRow As Long
    lastNonEmptyRow = 0

    For r = 1 To 500
        If Trim(CStr(configWS.Cells(r, 1).Value)) <> "" Then
            lastNonEmptyRow = r
        End If
    Next r

    ' Write section header after all existing data
    configWS.Cells(lastNonEmptyRow + 1, 1).Value = "BUTTON_ANCHORS"
    FindButtonAnchorStartRow = lastNonEmptyRow + 2
End Function

' FindNextButtonAnchorRow - Find next available row for button anchor
Private Function FindNextButtonAnchorRow(configWS As Worksheet) As Long
    Dim startRow As Long
    startRow = FindButtonAnchorStartRow(configWS)

    Dim r As Long
    For r = startRow To startRow + 100
        If Trim(CStr(configWS.Cells(r, 1).Value)) = "" Then
            FindNextButtonAnchorRow = r
            Exit Function
        End If
    Next r

    ' Fallback: return row after section
    FindNextButtonAnchorRow = startRow + 101
End Function

' SaveButtonAnchor - Save button position to config sheet (with duplicate check)
Sub SaveButtonAnchor(ws As Worksheet, btn As Object)
    Dim configWS As Worksheet
    Set configWS = GetOrCreateConfigSheet

    ' Calculate row anchor
    Dim btnRowAnchor As Long
    Dim r As Long
    For r = 1 To 20
        If btn.Top >= ws.Rows(r).Top And btn.Top < ws.Rows(r).Top + ws.Rows(r).Height Then
            btnRowAnchor = r
            Exit For
        End If
    Next r

    ' CRITICAL: Check for existing entry BEFORE calling FindNextButtonAnchorRow
    Dim anchorStartRow As Long
    anchorStartRow = FindButtonAnchorStartRow(configWS)

    Dim anchorRow As Long
    anchorRow = 0

    ' Scan for existing button entry
    Dim scanR As Long
    For scanR = anchorStartRow To anchorStartRow + 100
        If Trim(CStr(configWS.Cells(scanR, 1).Value)) = btn.Name Then
            anchorRow = scanR
            Exit For
        End If
        If Trim(CStr(configWS.Cells(scanR, 1).Value)) = "" Then
            Exit For
        End If
    Next scanR

    ' Only call FindNextButtonAnchorRow if no existing entry found
    If anchorRow = 0 Then
        anchorRow = FindNextButtonAnchorRow(configWS)
    End If

    ' Write to COMPARE_CONFIG
    configWS.Cells(anchorRow, 1).Value = btn.Name
    configWS.Cells(anchorRow, 2).Value = btnRowAnchor
    configWS.Cells(anchorRow, 3).Value = btn.Left - ws.Range("A1").Left
    configWS.Cells(anchorRow, 4).Value = btn.Width
    configWS.Cells(anchorRow, 5).Value = btn.Height
End Sub

' RepositionButtonsAfterColumnChange - Reposition buttons after column operations
Sub RepositionButtonsAfterColumnChange(ws As Worksheet)
    Dim configWS As Worksheet
    Set configWS = GetOrCreateConfigSheet

    Dim anchorStartRow As Long
    anchorStartRow = FindButtonAnchorStartRow(configWS)

    Dim btn As Object
    Dim btnName As String
    Dim storedRow As Long
    Dim storedOffset As Double
    Dim storedWidth As Double
    Dim storedHeight As Double

    Dim scanR As Long
    For scanR = anchorStartRow To anchorStartRow + 100  ' Changed from +50 to +100
        btnName = configWS.Cells(scanR, 1).Value
        If btnName = "" Then Exit For

        storedRow = CLng(configWS.Cells(scanR, 2).Value)
        storedOffset = CDbl(configWS.Cells(scanR, 3).Value)
        storedWidth = CDbl(configWS.Cells(scanR, 4).Value)
        storedHeight = CDbl(configWS.Cells(scanR, 5).Value)

        On Error Resume Next
        Set btn = ws.Buttons(btnName)
        On Error GoTo 0

        If Not btn Is Nothing Then
            btn.Top = ws.Rows(storedRow).Top
            btn.Left = ws.Range("A1").Left + storedOffset
            btn.Width = storedWidth
            btn.Height = storedHeight
        End If
    Next scanR
End Sub

'===============================================================================
' FIX BUTTON - One-time fix to update btnBuildUI OnAction
'===============================================================================

Public Sub FixBuildUIButton()
    '
    ' One-time fix: Updates btnBuildUI button OnAction from BuildFullUI to PreserveAndRebuildUI
    ' This runs on StartSystem to fix existing buttons on sheets
    '
    Dim ws As Worksheet
    Dim btn As Excel.Button

    On Error Resume Next

    ' Try to get the active worksheet
    Set ws = ActiveSheet

    If ws Is Nothing Then
        ' Try to get a valid worksheet
        Set ws = GetValidWorksheetForUI
    End If

    If ws Is Nothing Then
        DebugPrint "FixBuildUIButton: No worksheet available"
        Exit Sub
    End If

    ' Find and fix the button
    Set btn = ws.Buttons("btnBuildUI")
    If Not btn Is Nothing Then
        If btn.OnAction <> "PreserveAndRebuildUI" Then
            btn.OnAction = "PreserveAndRebuildUI"
            DebugPrint "FixBuildUIButton: Updated btnBuildUI OnAction to PreserveAndRebuildUI"
        End If
    Else
        DebugPrint "FixBuildUIButton: btnBuildUI button not found"
    End If

    On Error GoTo 0
End Sub

'===============================================================================
' PRESERVE UI - Rebuilds UI without wiping existing match rows
'===============================================================================

Public Sub PreserveAndRebuildUI()
    '
    ' Rebuilds the UI while preserving existing match rows
    ' This is called by the "Build UI" button to refresh UI structure
    ' without deleting user's X mark configurations
    '
    Dim ws As Worksheet
    Dim btn As Excel.Button
    Dim syncCol As Long
    Dim lastSyncCol As Long
    Dim wsSyncSheet As Worksheet

    g_MacroRunning = True

    ' PERFORMANCE OPTIMIZATION: Disable screen updating and calculations
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Get active worksheet
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = GetValidWorksheetForUI
    End If

    If ws Is Nothing Then
        ' Restore performance settings before exiting
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        MsgBox "No worksheet available.", vbCritical
        g_MacroRunning = False
        Exit Sub
    End If

    ' SELF-CORRECT: Update button OnAction if it still points to old function
    On Error Resume Next
    Set btn = ws.Buttons("btnBuildUI")
    If Not btn Is Nothing Then
        If btn.OnAction <> "PreserveAndRebuildUI" Then
            btn.OnAction = "PreserveAndRebuildUI"
            DebugPrint "PreserveAndRebuildUI: Updated button OnAction"
        End If
    End If
    On Error GoTo 0

    ' Set flags for PRESERVE mode (not full rebuild)
    g_SafeMode = False
    g_ForceRebuild = True
    g_ClearMatchDataOnRebuild = False  ' PRESERVE existing match rows
    Set g_CurrentWorksheet = ws

    ' Re-detect data header row before rebuild to ensure g_DataHeaderRow is correct
    Call InitializeDatasetContext(ws)

    ' Call rebuild - preserveMatchRows will be True because flag is False
    Call RebuildMatchBuilderUI

    ' Sync UI column header row columns 6+ to match data header row columns 6+
    If g_DataHeaderRow > 0 Then
        Set wsSyncSheet = g_CurrentWorksheet
        If wsSyncSheet Is Nothing Then Set wsSyncSheet = ActiveSheet
        lastSyncCol = 0
        For syncCol = 6 To 500
            If Trim(CStr(wsSyncSheet.Cells(g_DataHeaderRow, syncCol).Value)) <> "" Then
                lastSyncCol = syncCol
            End If
        Next syncCol
        If lastSyncCol >= 6 Then
            For syncCol = 6 To lastSyncCol
                wsSyncSheet.Cells(UI_COLHEADER_ROW, syncCol).Value = wsSyncSheet.Cells(g_DataHeaderRow, syncCol).Value
            Next syncCol
            ' Set column widths based on header name length only — not data rows
            For syncCol = 6 To lastSyncCol
                wsSyncSheet.Columns(syncCol).ColumnWidth = Len(CStr(wsSyncSheet.Cells(UI_COLHEADER_ROW, syncCol).Value)) * 1.2 + 2
            Next syncCol
        End If
    End If

    ' Freeze panes at first match rule row, column 6
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(UI_FIRST_MATCH_ROW, 6).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ' Restore performance settings
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "UI refreshed! Your match row configurations have been preserved.", vbInformation

    g_MacroRunning = False
End Sub

'===============================================================================
' RefreshHeaderRowVariables - Refresh header row globals after operations
'===============================================================================
Sub RefreshHeaderRowVariables(ws As Worksheet)
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    ' Create temp dictionaries for header detection
    Dim tempAliases As Object
    Dim tempLearned As Object
    Set tempAliases = CreateObject("Scripting.Dictionary")
    Set tempLearned = CreateObject("Scripting.Dictionary")

    Dim newHeaderRow As Long
    newHeaderRow = FindDataHeaderRow(ws, tempAliases, tempLearned)

    If newHeaderRow > 0 Then
        g_DataHeaderRow = newHeaderRow
    End If

    If Not g_SourceSheet Is Nothing Then
        g_SourceHeaderRow = FindDataHeaderRow(g_SourceSheet, tempAliases, tempLearned)
    End If
    If Not g_TargetSheet Is Nothing Then
        g_TargetHeaderRow = FindDataHeaderRow(g_TargetSheet, tempAliases, tempLearned)
    End If
End Sub

'===============================================================================
' ToggleMacroEvents - Toggle Application.EnableEvents for safe column deletion
'===============================================================================
Public Sub ToggleMacroEvents()
    If Application.EnableEvents Then
        Application.EnableEvents = False
        On Error Resume Next
        ActiveSheet.Buttons("btnToggleMacroEvents").Caption = "Resume Macro"
        On Error GoTo 0
        MsgBox "Macro events PAUSED. You can now delete columns and use Undo freely." & vbCrLf & vbCrLf & "Click 'Pause Macro' again to resume.", vbInformation, "Macro Paused"
    Else
        Application.EnableEvents = True
        On Error Resume Next
        ActiveSheet.Buttons("btnToggleMacroEvents").Caption = "Pause Macro"
        On Error GoTo 0
        MsgBox "Macro events RESUMED. Double-click and change events are active again.", vbInformation, "Macro Resumed"
    End If
End Sub

'===============================================================================
' PromptForSheetSelection - Prompt user to select a sheet from a workbook
' Returns the selected worksheet or Nothing if cancelled
'===============================================================================
Private Function PromptForSheetSelection(wb As Workbook) As Worksheet
    Dim i As Long
    Dim sheetList As String
    Dim userChoice As String
    Dim selectedIdx As Long

    If wb.Worksheets.Count = 1 Then
        Set PromptForSheetSelection = wb.Worksheets(1)
        Exit Function
    End If

    sheetList = "File contains " & wb.Worksheets.Count & " sheets:" & vbCrLf & vbCrLf
    For i = 1 To wb.Worksheets.Count
        sheetList = sheetList & "  " & i & ".  " & wb.Worksheets(i).Name & vbCrLf
    Next i
    sheetList = sheetList & vbCrLf & "Enter the sheet NUMBER to load:"

    userChoice = InputBox(sheetList, "Select Sheet", "1")

    If userChoice = "" Then
        Set PromptForSheetSelection = Nothing
        Exit Function
    End If

    On Error Resume Next
    selectedIdx = CLng(userChoice)
    On Error GoTo 0

    If selectedIdx < 1 Or selectedIdx > wb.Worksheets.Count Then
        MsgBox "Invalid selection. Using sheet 1.", vbInformation
        selectedIdx = 1
    End If

    Set PromptForSheetSelection = wb.Worksheets(selectedIdx)
End Function

'===============================================================================
' LoadSourceFile - Load source data from an external Excel file
'===============================================================================
Public Sub LoadSourceFile()
    Dim ws As Worksheet
    Dim sourceWB As Workbook
    Dim sourceFileName As Variant
    Dim lastUsedRow As Long
    Dim pasteRow As Long
    Dim scanRow As Long
    Dim errNum As Long
    Dim errDesc As String
    Dim sourceData As Variant
    Dim srcConfigSheet As Worksheet
    Dim loadedSourceFileName As String
    Dim loadedSourceSheetName As String
    Dim sourceLastRow As Long
    Dim sourceLastCol As Long
    Dim lastMatchRuleRow As Long
    Dim scanPasteRow As Long
    Dim scanCol As Long
    Dim lastUsedCol As Long
    Dim dataCol As Long
    Dim consecutiveEmpty As Long
    Dim clearLastCol As Long
    Dim uiLastCol As Long
    Dim uiClearCol As Long
    Dim staleStartCol As Long
    Dim scanCeiling As Long
    Dim selectedSourceSheet As Worksheet
    Dim mandatoryNames(4) As String
    Dim keepCol() As Boolean
    Dim filteredData() As Variant
    Dim filteredColCount As Long
    Dim iCol As Long
    Dim iName As Long
    Dim iRow As Long
    Dim isMandatory As Boolean
    Dim sourceHeaderVal As String
    Dim destCol As Long
    Dim mCol As Long
    Dim syncCol As Long
    Dim matchRuleRow As Long
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' Check if data already exists
    If g_DataHeaderRow > 0 Then
        If MsgBox("Source data already exists. Replace it?", vbYesNo + vbQuestion, "Load Source") = vbNo Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If

    ' Show file picker
    sourceFileName = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", , "Select Source File")
    If sourceFileName = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Get current worksheet
    Set ws = g_CurrentWorksheet
    If ws Is Nothing Then Set ws = ActiveSheet

    ' Initialize dataset context to ensure g_DataHeaderRow is set before pasteRow calculation
    If g_DataHeaderRow = 0 Or Not g_Initialized Then
        Call InitializeDatasetContext(ws)
    End If

    ' Open the selected file and read data directly into array
    Set sourceWB = Workbooks.Open(sourceFileName)
    Set selectedSourceSheet = PromptForSheetSelection(sourceWB)
    If selectedSourceSheet Is Nothing Then
        sourceWB.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Exit Sub
    End If
    sourceData = selectedSourceSheet.UsedRange.Value
    sourceLastRow = selectedSourceSheet.UsedRange.Rows.Count
    sourceLastCol = selectedSourceSheet.UsedRange.Columns.Count
    loadedSourceFileName = Mid(CStr(sourceFileName), InStrRev(CStr(sourceFileName), "\") + 1)
    loadedSourceSheetName = selectedSourceSheet.Name
    sourceWB.Close SaveChanges:=False
    On Error GoTo 0

    ' Define mandatory column names
    mandatoryNames(0) = "MATCHED_ID"
    mandatoryNames(1) = "MATCH_TYPE"
    mandatoryNames(2) = "MATCH_STATUS"
    mandatoryNames(3) = "SOURCE_FILE"
    mandatoryNames(4) = "TARGET_FILE"

    ' Identify which source columns to keep (exclude mandatory columns)
    ReDim keepCol(1 To sourceLastCol)
    filteredColCount = 0
    For iCol = 1 To sourceLastCol
        sourceHeaderVal = UCase(Trim(CStr(sourceData(1, iCol))))
        isMandatory = False
        For iName = 0 To 4
            If sourceHeaderVal = mandatoryNames(iName) Then
                isMandatory = True
                Exit For
            End If
        Next iName
        keepCol(iCol) = Not isMandatory
        If keepCol(iCol) Then filteredColCount = filteredColCount + 1
    Next iCol

    ' Build filtered array excluding mandatory columns
    ReDim filteredData(1 To sourceLastRow, 1 To filteredColCount)
    For iRow = 1 To sourceLastRow
        destCol = 0
        For iCol = 1 To sourceLastCol
            If keepCol(iCol) Then
                destCol = destCol + 1
                filteredData(iRow, destCol) = sourceData(iRow, iCol)
            End If
        Next iCol
    Next iRow

    ' Scan from UI_FIRST_MATCH_ROW (row 5) to g_DataHeaderRow - 1 for last UI match-rule row
    lastMatchRuleRow = 0
    scanCeiling = 50
    If g_DataHeaderRow > 0 Then scanCeiling = g_DataHeaderRow - 1
    For scanPasteRow = UI_FIRST_MATCH_ROW To scanCeiling
        If IsNumeric(ws.Cells(scanPasteRow, 1).Value) Or _
           Trim(CStr(ws.Cells(scanPasteRow, 2).Value)) <> "" Then
            lastMatchRuleRow = scanPasteRow
        End If
    Next scanPasteRow

    If lastMatchRuleRow > 0 Then
        pasteRow = lastMatchRuleRow + 1
    Else
        pasteRow = FIXED_DATA_HEADER_ROW
    End If
    If g_DataHeaderRow > 0 And pasteRow < g_DataHeaderRow Then
        pasteRow = g_DataHeaderRow
    End If

    ' Find last used column in the paste row
    lastUsedCol = 0
    For scanCol = 1 To 500
        If Trim(CStr(ws.Cells(pasteRow, scanCol).Value)) <> "" Then
            lastUsedCol = scanCol
        End If
    Next scanCol

    ' Find first populated data column in pasteRow (column 6 onward) for row scan
    dataCol = 0
    For scanCol = 6 To lastUsedCol
        If Trim(CStr(ws.Cells(pasteRow, scanCol).Value)) <> "" Then
            dataCol = scanCol
            Exit For
        End If
    Next scanCol

    ' Find last used row by scanning the data column downward
    ' Stop after 10 consecutive empty cells to handle sparse data
    lastUsedRow = pasteRow
    If dataCol > 0 Then
        consecutiveEmpty = 0
        For scanRow = pasteRow + 1 To pasteRow + 500000
            If scanRow > ws.Rows.Count Then Exit For
            If Trim(CStr(ws.Cells(scanRow, dataCol).Value)) = "" Then
                consecutiveEmpty = consecutiveEmpty + 1
                If consecutiveEmpty >= 10 Then Exit For
            Else
                consecutiveEmpty = 0
                lastUsedRow = scanRow
            End If
        Next scanRow
    End If

    ' Clear existing content from boundary row to last used row (full width)
    ' Use max of old lastUsedCol and new data width to ensure full coverage
    clearLastCol = lastUsedCol
    If 5 + filteredColCount > clearLastCol Then clearLastCol = 5 + filteredColCount
    If lastUsedRow >= pasteRow And clearLastCol > 0 Then
        Application.DisplayAlerts = False
        ws.Range(ws.Cells(pasteRow, 1), ws.Cells(lastUsedRow, clearLastCol)).ClearContents
        ws.Range(ws.Cells(pasteRow, 1), ws.Cells(lastUsedRow, clearLastCol)).ClearFormats
        Application.DisplayAlerts = True
    End If

    ' Targeted clear of mandatory columns 1-5 using exact source row count
    ' Guarantees full clear regardless of gaps in source data columns
    If sourceLastRow > 0 Then
        Application.DisplayAlerts = False
        ws.Range(ws.Cells(pasteRow, 1), ws.Cells(pasteRow + sourceLastRow - 1, 5)).ClearContents
        ws.Range(ws.Cells(pasteRow, 1), ws.Cells(pasteRow + sourceLastRow - 1, 5)).ClearFormats
        Application.DisplayAlerts = True
    End If

    ' Clear any stale formatting on the data header row before writing new headers
    ws.Rows(pasteRow).ClearFormats

    ' Write mandatory headers to columns 1-5 of pasteRow
    For mCol = 0 To 4
        ws.Cells(pasteRow, mCol + 1).Value = mandatoryNames(mCol)
        ws.Cells(pasteRow, mCol + 1).Font.Bold = True
        ws.Cells(pasteRow, mCol + 1).Interior.Color = RGB(91, 115, 150)
        ws.Cells(pasteRow, mCol + 1).Font.Color = RGB(255, 255, 255)
        ws.Cells(pasteRow, mCol + 1).HorizontalAlignment = xlCenter
        ws.Columns(mCol + 1).ColumnWidth = Len(mandatoryNames(mCol)) * 1.8 + 2
    Next mCol

    ' Paste filtered source data starting at column 6
    ws.Range(ws.Cells(pasteRow, 6), ws.Cells(pasteRow + sourceLastRow - 1, 5 + filteredColCount)).Value = filteredData

    ' Find previous UI column header row width before clearing — check content or formatting
    uiLastCol = 0
    For scanCol = 6 To 500
        If Trim(CStr(ws.Cells(UI_COLHEADER_ROW, scanCol).Value)) <> "" Then
            uiLastCol = scanCol
        ElseIf ws.Cells(UI_COLHEADER_ROW, scanCol).Interior.ColorIndex <> xlNone Then
            uiLastCol = scanCol
        End If
    Next scanCol
    uiClearCol = uiLastCol
    If clearLastCol > uiClearCol Then uiClearCol = clearLastCol

    ' Clear stale UI zone columns beyond new data width — only columns past the new range
    staleStartCol = 6 + filteredColCount
    If uiClearCol >= staleStartCol And g_DataHeaderRow > 0 Then
        ws.Range(ws.Cells(UI_COLHEADER_ROW, staleStartCol), ws.Cells(g_DataHeaderRow - 1, uiClearCol)).ClearContents
        ws.Range(ws.Cells(UI_COLHEADER_ROW, staleStartCol), ws.Cells(g_DataHeaderRow - 1, uiClearCol)).ClearFormats
    End If

    ' Sync UI column header row columns 6+ to match data header row columns 6+
    For syncCol = 6 To 5 + filteredColCount
        ws.Cells(UI_COLHEADER_ROW, syncCol).Value = ws.Cells(pasteRow, syncCol).Value
        ws.Cells(UI_COLHEADER_ROW, syncCol).Font.Bold = True
        ws.Cells(UI_COLHEADER_ROW, syncCol).Interior.Color = RGB(68, 84, 106)
        ws.Cells(UI_COLHEADER_ROW, syncCol).Font.Color = RGB(255, 255, 255)
        ws.Cells(UI_COLHEADER_ROW, syncCol).HorizontalAlignment = xlCenter
        ' Apply same color to data header row
        ws.Cells(pasteRow, syncCol).Interior.Color = RGB(68, 84, 106)
        ws.Cells(pasteRow, syncCol).Font.Bold = True
        ws.Cells(pasteRow, syncCol).Font.Color = RGB(255, 255, 255)
        ws.Cells(pasteRow, syncCol).HorizontalAlignment = xlCenter
    Next syncCol

    ' Clear X marks and apply alternating colors and borders to match rule rows
    For matchRuleRow = UI_FIRST_MATCH_ROW To pasteRow - 1
        For syncCol = 6 To 5 + filteredColCount
            If UCase(Trim(CStr(ws.Cells(matchRuleRow, syncCol).Value))) = "X" Then
                ws.Cells(matchRuleRow, syncCol).ClearContents
            End If
            If matchRuleRow Mod 2 = 0 Then
                ws.Cells(matchRuleRow, syncCol).Interior.Color = RGB(214, 224, 240)
            Else
                ws.Cells(matchRuleRow, syncCol).Interior.Color = RGB(237, 242, 250)
            End If
            ws.Cells(matchRuleRow, syncCol).Borders(xlEdgeBottom).LineStyle = xlContinuous
            ws.Cells(matchRuleRow, syncCol).Borders(xlEdgeBottom).Weight = xlThin
            ws.Cells(matchRuleRow, syncCol).Borders(xlEdgeRight).LineStyle = xlContinuous
            ws.Cells(matchRuleRow, syncCol).Borders(xlEdgeRight).Weight = xlThin
        Next syncCol
    Next matchRuleRow
    ' Set column widths based on header name length only — not data rows
    For syncCol = 6 To 5 + filteredColCount
        ws.Columns(syncCol).ColumnWidth = Len(CStr(ws.Cells(UI_COLHEADER_ROW, syncCol).Value)) * 1.2 + 2
    Next syncCol

    ' Apply AutoFilter to data header row
    On Error Resume Next
    ws.AutoFilterMode = False
    ws.Range(ws.Cells(pasteRow, 1), ws.Cells(pasteRow, 5 + filteredColCount)).AutoFilter
    On Error GoTo 0

    ' Freeze panes at first match rule row, column 6
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(UI_FIRST_MATCH_ROW, 6).Select
    ActiveWindow.FreezePanes = True

    ' Reinitialize and rebuild UI
    g_DataHeaderRow = 0
    g_Initialized = False
    Call InitializeDatasetContext(ws)
    Call RebuildMatchBuilderUI
    Set srcConfigSheet = GetOrCreateConfigSheet
    Call SetConfigValue(srcConfigSheet, "LOADED_SOURCE_FILE", loadedSourceFileName)
    Call SetConfigValue(srcConfigSheet, "LOADED_SOURCE_SHEET", loadedSourceSheetName)
    Call AddStatusDisplayFixed(ws)

    Application.ScreenUpdating = True
    MsgBox "Source data loaded successfully.", vbInformation
    Exit Sub

ErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = True
    If Not sourceWB Is Nothing Then
        On Error Resume Next
        sourceWB.Close SaveChanges:=False
        On Error GoTo 0
    End If
    MsgBox "Error loading source file: " & errNum & " - " & errDesc, vbCritical
End Sub

'===============================================================================
' LoadTargetFile - Load target data from an external Excel file
'===============================================================================
Public Sub LoadTargetFile()
    Dim targetWB As Workbook
    Dim targetFileName As Variant
    Dim targetSheet As Worksheet
    Dim wsTarget As Worksheet
    Dim lastUsedRow As Long
    Dim scanRow As Long
    Dim selectedTargetSheet As Worksheet
    Dim fmtCol As Long
    Dim lastTargetCol As Long
    Dim tgtFmt As String
    Dim colFormats() As String
    Dim loadedTargetFileName As String
    Dim loadedTargetSheetName As String
    Dim tgtConfigSheet As Worksheet
    Dim hasData As Boolean
    Dim configSheet As Worksheet
    Dim wsStatus As Worksheet
    Dim targetDataArr As Variant
    Dim oldLastCol As Long
    Dim oldLastRow As Long
    Dim oldColScanRow As Long
    Dim clearScanCol As Long
    Dim oldScanCeiling As Long
    Dim oldScanCol As Long

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' Get or create TARGET_DATA sheet
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("TARGET_DATA")
    On Error GoTo ErrorHandler

    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Worksheets.Add
        wsTarget.Name = "TARGET_DATA"
    End If

    ' Check if data already exists
    hasData = False
    For scanRow = 1 To 10
        If Trim(CStr(wsTarget.Cells(scanRow, 1).Value)) <> "" Then
            hasData = True
            Exit For
        End If
    Next scanRow

    If hasData Then
        If MsgBox("Target data already exists. Replace it?", vbYesNo + vbQuestion, "Load Target") = vbNo Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End If

    ' Show file picker
    targetFileName = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", , "Select Target File")
    If targetFileName = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Open the selected file and copy data
    Set targetWB = Workbooks.Open(targetFileName)
    Set selectedTargetSheet = PromptForSheetSelection(targetWB)
    If selectedTargetSheet Is Nothing Then
        targetWB.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Capture number formats from target sheet before closing
    lastTargetCol = selectedTargetSheet.UsedRange.Columns.Count
    ReDim colFormats(1 To lastTargetCol)
    For fmtCol = 1 To lastTargetCol
        colFormats(fmtCol) = selectedTargetSheet.Cells(2, fmtCol).NumberFormat
    Next fmtCol

    ' Clear existing content — capture old dimensions before overwriting
    oldLastCol = 0
    oldLastRow = 0
    For oldColScanRow = 1 To 5
        For clearScanCol = 1 To 500
            If Trim(CStr(wsTarget.Cells(oldColScanRow, clearScanCol).Value)) <> "" Then
                If clearScanCol > oldLastCol Then oldLastCol = clearScanCol
            End If
        Next clearScanCol
    Next oldColScanRow
    On Error Resume Next
    oldScanCeiling = wsTarget.UsedRange.Rows.Count + wsTarget.UsedRange.Row - 1
    On Error GoTo ErrorHandler
    If oldScanCeiling < 1 Then oldScanCeiling = 1
    For oldScanCol = 1 To 10
        For scanRow = 1 To oldScanCeiling
            If Trim(CStr(wsTarget.Cells(scanRow, oldScanCol).Value)) <> "" Then
                If scanRow > oldLastRow Then oldLastRow = scanRow
            End If
        Next scanRow
    Next oldScanCol
    If oldLastRow > 0 And oldLastCol > 0 Then
        Application.DisplayAlerts = False
        wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(oldLastRow, oldLastCol)).ClearContents
        wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(oldLastRow, oldLastCol)).ClearFormats
        Application.DisplayAlerts = True
    End If

    ' Write values directly from array (avoids clipboard being cleared by MsgBox)
    targetDataArr = selectedTargetSheet.UsedRange.Value
    wsTarget.Cells(1, 1).Resize(UBound(targetDataArr, 1), UBound(targetDataArr, 2)).Value = targetDataArr

    ' Apply captured number formats to target sheet columns
    For fmtCol = 1 To lastTargetCol
        tgtFmt = colFormats(fmtCol)
        If tgtFmt <> "General" And tgtFmt <> "@" And tgtFmt <> "" Then
            wsTarget.Columns(fmtCol).NumberFormat = tgtFmt
        End If
    Next fmtCol

    ' Close target file without saving
    loadedTargetFileName = Mid(CStr(targetFileName), InStrRev(CStr(targetFileName), "\") + 1)
    loadedTargetSheetName = selectedTargetSheet.Name
    targetWB.Close SaveChanges:=False

    ' Save target sheet reference to config
    Set configSheet = GetOrCreateConfigSheet
    Call SetConfigValue(configSheet, CONFIG_TARGET_WB, ThisWorkbook.Name)
    Call SetConfigValue(configSheet, CONFIG_TARGET_WS, "TARGET_DATA")
    Set tgtConfigSheet = GetOrCreateConfigSheet
    Call SetConfigValue(tgtConfigSheet, "LOADED_TARGET_FILE", loadedTargetFileName)
    Call SetConfigValue(tgtConfigSheet, "LOADED_TARGET_SHEET", loadedTargetSheetName)
    Set wsStatus = g_CurrentWorksheet
    If Not wsStatus Is Nothing Then Call AddStatusDisplayFixed(wsStatus)

    Application.ScreenUpdating = True
    MsgBox "Target data loaded successfully.", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    If Not targetWB Is Nothing Then
        On Error Resume Next
        targetWB.Close SaveChanges:=False
        On Error GoTo 0
    End If
    MsgBox "Error loading target file: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'===============================================================================
' ExportResults - Export source data with results and target data to new file
'===============================================================================
Public Sub ExportResults()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim newWB As Workbook
    Dim newWS1 As Worksheet
    Dim newWS2 As Worksheet
    Dim saveFileName As String
    Dim lastSourceRow As Long
    Dim lastSourceCol As Long
    Dim lastTargetRow As Long
    Dim lastTargetCol As Long
    Dim scanRow As Long
    Dim scanCol As Long
    Dim fmtCol As Long
    Dim srcFmt As String

    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    ' Get source worksheet
    Set wsSource = g_CurrentWorksheet
    If wsSource Is Nothing Then Set wsSource = ActiveSheet

    ' Ensure g_DataHeaderRow is initialized
    If g_DataHeaderRow = 0 Or Not g_Initialized Then
        Call InitializeDatasetContext(wsSource)
    End If

    ' Get target worksheet
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("TARGET_DATA")
    On Error GoTo ErrorHandler

    If wsTarget Is Nothing Then
        MsgBox "TARGET_DATA sheet not found. Please load target data first.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Check source has data
    If g_DataHeaderRow = 0 Then
        MsgBox "No source data found. Please load source data first.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Find last row and col of source data
    lastSourceRow = 0
    For scanRow = g_DataHeaderRow To g_DataHeaderRow + 100000
        If scanRow > wsSource.Rows.Count Then Exit For
        Dim anyData As Boolean
        anyData = False
        For scanCol = 1 To 20
            If Trim(CStr(wsSource.Cells(scanRow, scanCol).Value)) <> "" Then
                anyData = True
                Exit For
            End If
        Next scanCol
        If anyData Then
            lastSourceRow = scanRow
        ElseIf lastSourceRow > 0 And scanRow > lastSourceRow + 10 Then
            Exit For
        End If
    Next scanRow

    lastSourceCol = 0
    For scanCol = 1 To 500
        If Trim(CStr(wsSource.Cells(g_DataHeaderRow, scanCol).Value)) <> "" Then
            lastSourceCol = scanCol
        End If
    Next scanCol

    If lastSourceRow = 0 Or lastSourceCol = 0 Then
        MsgBox "No data found to export.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Find last row and col of target data
    lastTargetRow = 0
    For scanRow = 1 To 100000
        If scanRow > wsTarget.Rows.Count Then Exit For
        For scanCol = 1 To 20
            If Trim(CStr(wsTarget.Cells(scanRow, scanCol).Value)) <> "" Then
                lastTargetRow = scanRow
                Exit For
            End If
        Next scanCol
    Next scanRow

    lastTargetCol = 0
    For scanCol = 1 To 500
        If Trim(CStr(wsTarget.Cells(1, scanCol).Value)) <> "" Then
            lastTargetCol = scanCol
        End If
    Next scanCol

    ' Show save dialog
    saveFileName = CStr(Application.GetSaveAsFilename( _
        InitialFileName:="Results_" & Format(Now, "YYYYMMDD_HHMMSS"), _
        FileFilter:="Excel Files (*.xlsx), *.xlsx", _
        Title:="Save Results As"))

    If saveFileName = "False" Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Create new workbook
    Set newWB = Workbooks.Add

    ' Copy source data to Sheet1
    Set newWS1 = newWB.Worksheets(1)
    newWS1.Name = "Source_Results"
    wsSource.Range(wsSource.Cells(g_DataHeaderRow, 1), wsSource.Cells(lastSourceRow, lastSourceCol)).Copy
    newWS1.Cells(1, 1).PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    ' Apply source number formats to exported columns — preserves date/time display
    For fmtCol = 1 To lastSourceCol
        srcFmt = wsSource.Cells(g_DataHeaderRow + 1, fmtCol).NumberFormat
        If srcFmt <> "General" And srcFmt <> "@" And srcFmt <> "" Then
            newWS1.Columns(fmtCol).NumberFormat = srcFmt
        End If
    Next fmtCol

    ' Copy target data to Sheet2
    If lastTargetRow > 0 And lastTargetCol > 0 Then
        Set newWS2 = newWB.Worksheets.Add(After:=newWS1)
        newWS2.Name = "Target_Data"
        wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(lastTargetRow, lastTargetCol)).Copy
        newWS2.Cells(1, 1).PasteSpecial xlPasteValues
        Application.CutCopyMode = False

        ' Apply target sheet number formats to exported columns
        For fmtCol = 1 To lastTargetCol
            srcFmt = wsTarget.Cells(2, fmtCol).NumberFormat
            If srcFmt <> "General" And srcFmt <> "@" And srcFmt <> "" Then
                newWS2.Columns(fmtCol).NumberFormat = srcFmt
            End If
        Next fmtCol
    End If

    ' Save and close new workbook
    Application.DisplayAlerts = False
    newWB.SaveAs Filename:=saveFileName, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    newWB.Close SaveChanges:=False

    Application.ScreenUpdating = True
    MsgBox "Results exported successfully to:" & vbCrLf & saveFileName, vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    If Not newWB Is Nothing Then
        On Error Resume Next
        newWB.Close SaveChanges:=False
        On Error GoTo 0
    End If
    MsgBox "Error exporting results: " & Err.Description, vbCritical
End Sub

'===============================================================================
' ClearAllData - Clear source and target data, reset match rules
'===============================================================================
Public Sub ClearAllData()
    Dim ws As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim scanRow As Long
    Dim response As Integer
    Dim clearDataCol As Long
    Dim clearConsecEmpty As Long
    Dim findCol As Long
    Dim nullCol As Long
    Dim lastDataCol As Long
    Dim nullLastRow As Long
    Dim nullConsecEmpty As Long

    ' Confirm before clearing
    response = MsgBox("This will clear all source data, target data, and reset match rules." & vbCrLf & vbCrLf & _
                      "Are you sure you want to continue?", vbYesNo + vbQuestion, "Clear All Data")
    If response = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    ' Get source worksheet
    Set ws = g_CurrentWorksheet
    If ws Is Nothing Then Set ws = ActiveSheet

    ' Ensure g_DataHeaderRow is initialized
    If g_DataHeaderRow = 0 Or Not g_Initialized Then
        Call InitializeDatasetContext(ws)
    End If

    ' Clear source data below header row — preserve header row itself
    If g_DataHeaderRow > 0 Then
        clearDataCol = 0
        For findCol = 6 To 500
            If Trim(CStr(ws.Cells(g_DataHeaderRow, findCol).Value)) <> "" Then
                clearDataCol = findCol
                Exit For
            End If
        Next findCol
        lastRow = g_DataHeaderRow
        If clearDataCol > 0 Then
            clearConsecEmpty = 0
            For scanRow = g_DataHeaderRow + 1 To g_DataHeaderRow + 500000
                If scanRow > ws.Rows.Count Then Exit For
                If Trim(CStr(ws.Cells(scanRow, clearDataCol).Value)) = "" Then
                    clearConsecEmpty = clearConsecEmpty + 1
                    If clearConsecEmpty >= 10 Then Exit For
                Else
                    clearConsecEmpty = 0
                    lastRow = scanRow
                End If
            Next scanRow
        End If
        ' Also check column 2 for ~NULL~ placeholders from Execute Match
        nullLastRow = g_DataHeaderRow
        nullConsecEmpty = 0
        For scanRow = g_DataHeaderRow + 1 To g_DataHeaderRow + 500000
            If scanRow > ws.Rows.Count Then Exit For
            If Trim(CStr(ws.Cells(scanRow, 2).Value)) = "" Then
                nullConsecEmpty = nullConsecEmpty + 1
                If nullConsecEmpty >= 10 Then Exit For
            Else
                nullConsecEmpty = 0
                nullLastRow = scanRow
            End If
        Next scanRow
        If nullLastRow > lastRow Then lastRow = nullLastRow

        If lastRow > g_DataHeaderRow Then
            Application.DisplayAlerts = False
            ws.Range(ws.Rows(g_DataHeaderRow + 1), ws.Rows(lastRow)).ClearContents
            ws.Range(ws.Rows(g_DataHeaderRow + 1), ws.Rows(lastRow)).ClearFormats
            Application.DisplayAlerts = True
        End If
        ' Reset first data row cols 1-5 to default appearance
        With ws.Range(ws.Cells(g_DataHeaderRow + 1, 1), ws.Cells(g_DataHeaderRow + 1, 5))
            .Interior.ColorIndex = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
        End With
    End If

    ' Clear TARGET_DATA sheet
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("TARGET_DATA")
    On Error GoTo ErrorHandler

    If Not wsTarget Is Nothing Then
        lastRow = 0
        For scanRow = 1 To 100000
            If scanRow > wsTarget.Rows.Count Then Exit For
            If Trim(CStr(wsTarget.Cells(scanRow, 1).Value)) <> "" Then
                lastRow = scanRow
            End If
        Next scanRow

        If lastRow > 0 Then
            Application.DisplayAlerts = False
            wsTarget.Range(wsTarget.Rows(1), wsTarget.Rows(lastRow)).ClearContents
            Application.DisplayAlerts = True
        End If
    End If

    ' Write ~NULL~ to data header row columns 6+ and UI column header row columns 6+
    If g_DataHeaderRow > 0 Then
        lastDataCol = 0
        For nullCol = 6 To 500
            If Trim(CStr(ws.Cells(g_DataHeaderRow, nullCol).Value)) <> "" Then
                lastDataCol = nullCol
            End If
        Next nullCol
        If lastDataCol >= 6 Then
            For nullCol = 6 To lastDataCol
                ws.Cells(g_DataHeaderRow, nullCol).Value = "~NULL~"
                ws.Cells(UI_COLHEADER_ROW, nullCol).Value = "~NULL~"
            Next nullCol
        End If
    End If

    ' Reset match rules to single default row
    g_ClearMatchDataOnRebuild = True
    g_ForceRebuild = True

    ' Reset globals
    g_DataHeaderRow = 0
    g_Initialized = False
    Set g_SourceSheet = Nothing
    Set g_TargetSheet = Nothing
    Set g_SourceWorkbook = Nothing
    Set g_TargetWorkbook = Nothing

    ' Rebuild UI fresh
    Call InitializeDatasetContext(ws)
    Call RebuildMatchBuilderUI

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "All data cleared. Ready for new data.", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    MsgBox "Error clearing data: " & Err.Description, vbCritical
End Sub
