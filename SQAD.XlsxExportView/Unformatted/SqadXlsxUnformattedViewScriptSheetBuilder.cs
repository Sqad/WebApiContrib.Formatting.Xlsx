using System.Data;
using System.Linq;
using OfficeOpenXml;
using SQAD.XlsxExportImport.Base.Builders;
using SQAD.XlsxExportView.Unformatted.Models;

namespace SQAD.XlsxExportView.Unformatted
{
    internal class SqadXlsxUnformattedViewScriptSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private const string RefreshDataQueryTableName = "RefreshDataQueryTable";

        private readonly string _dataSheetName;
        private readonly UnformattedExportSettings _settings;
        private readonly bool _needCreatePivotSheet;

        public SqadXlsxUnformattedViewScriptSheetBuilder(UnformattedExportSettings settings, bool needCreatePivotSheet
              , string dataSheetName = null)
            : base(ExportViewConstants.ScriptSheetName, shouldAutoFit: false)
        {
            _settings = settings;
            _needCreatePivotSheet = needCreatePivotSheet;
            _dataSheetName = string.IsNullOrEmpty(dataSheetName)
                 ? ExportViewConstants.UnformattedViewDataSheetName : dataSheetName;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (worksheet.Workbook.VbaProject == null)
            {
                worksheet.Workbook.CreateVBAProject();
            }

            var dataConstantsScript = string.Empty;
            var dataScript = string.Empty;
            var refreshDataScript = string.Empty;

            var dataSheet = worksheet.Workbook
                                     .Worksheets
                                     .First(x => x.Name == _dataSheetName);
            
            if (_settings != null)
            {
                dataConstantsScript = GetDataConstantsScript();
                dataScript = GetDataScript(dataSheet);
                refreshDataScript = 
                    //_settings.UseNewVersion && !_settings.UseEmbeddedLogin
                    //? GetBrowserRefreshDataScript()
                    //: 
                    GetQueryTableRefreshDataScript();
            }

            var pivotScript = string.Empty;
            if (_needCreatePivotSheet)
            {
                if (dataSheet.Dimension == null)
                {
                    dataSheet.Cells[1,1].Value = "";
                }
                pivotScript = GetPivotScript(dataSheet.Dimension.ToString());
            }

            var code = $@"
{dataConstantsScript}

Private Sub Workbook_Open()
    Dim tmpSheet As Worksheet
    Set tmpSheet = Sheets(""{ExportViewConstants.ScriptSheetName}"")
    Dim sheetData As Worksheet
    Set sheetData = Sheets(""{_dataSheetName}"")
    For n = sheetData.QueryTables.Count To 1 Step - 1
      sheetData.QueryTables(n).Delete
    Next n
    For n = ThisWorkbook.Connections.Count To 1 Step - 1
      ThisWorkbook.Connections.Item(n).Delete
    Next n

    InitQueryTable

    If tmpSheet.Visible <> xlSheetVeryHidden Then
        
      {dataScript}
    
      {pivotScript}

      tmpSheet.Visible = xlSheetVeryHidden
    End If
End Sub

Private Sub InitQueryTable()
    
    Dim sheet As Worksheet
    Set sheet = Sheets(""{_dataSheetName}"")


    Dim qt As QueryTable
    If sheet.QueryTables.Count = 0 Then
      Set qt = sheet.QueryTables.Add(Connection:= ""URL;"" & ExportUrl, Destination:= Sheets(""{_dataSheetName}"").Range(""A1""))
    Else
      Set qt = sheet.QueryTables(1)
    End If

    qt.AdjustColumnWidth = True
    qt.RefreshStyle = xlOverwriteCells
    qt.RefreshOnFileOpen = False
    qt.BackgroundQuery = False

    qt.Name = ""{RefreshDataQueryTableName}""
    qt.WebPreFormattedTextToColumns = True
    qt.WebFormatting = xlWebFormattingNone
    qt.WebConsecutiveDelimitersAsOne = True
    qt.WebDisableRedirections = False
    qt.WebSingleBlockTextImport = False
    qt.WebDisableDateRecognition = False
    qt.WebSelectionType = xlEntirePage
    isRefreshed = False

End Sub

{ refreshDataScript}
";
            worksheet.Workbook.CodeModule.Code = code;
        }

        private string GetDataConstantsScript()
        {
            var constants = $@"
    Private Const ExportUrl = ""{_settings.ExcelLink}""
    Private isRefreshed As Boolean
";

            if (!_settings.UseEmbeddedLogin)
            {
                constants = $@"
{constants}
    Private Const TokenPageLink = ""{_settings.TokenPageLink}""
    Private Const LoginPageLink = ""{_settings.LoginPageLink}""
";
            }

            return constants;
        }

        private string GetDataScript(ExcelWorksheet worksheet)
        {
            const string cellsName = "nwshp_hl_en_tab_wn";
            var cells = worksheet.Cells[worksheet.Dimension.Address];
            worksheet.Names.Add(cellsName, cells);

            return $@"
    Dim instrSheet As Worksheet
    Set instrSheet = Sheets(""{ExportViewConstants.UnformattedViewInstructionsSheetName}"")

    Dim position As Range
    Set position = instrSheet.Range(instrSheet.Cells(5, 2), instrSheet.Cells(5, 3))
    
    Dim btn As Button
    Set btn = instrSheet.Buttons.Add(position.Left, position.Top, position.Width, position.Height)
    With btn
        .Name = ""RefreshDataButton""
        .Caption = ""Refresh Connection""
        .OnAction = ""ThisWorkbook.RefreshButtonClick""
    End With
";
        }

        private string GetQueryTableRefreshDataScript()
        {
            return $@"
Sub RefreshButtonClick()
    If isRefreshed Then
      MsgBox ""Connection is refreshed.Use 'Edit Query' to update data""
      Exit Sub
    End If

    Dim sheet As Worksheet
    Dim tempSheet As Worksheet
    Dim qt As QueryTable

    Set sheet = Sheets(""{_dataSheetName}"")
    Set qt = sheet.QueryTables(1)
    Sheets.Add.Name = ""TempData""
    Set tempSheet = Sheets(""TempData"")
    tempSheet.Visible = xlSheetHidden
    Sheets(""{_dataSheetName}"").Range(sheet.UsedRange.Address).Copy (Sheets(""TempData"").Range(sheet.UsedRange.Address))
    sheet.Cells.Clear
    On Error Resume Next
    qt.Refresh
    sheet.Cells.Clear
    Sheets(""TempData"").Range(tempSheet.UsedRange.Address).Copy (Sheets(""{_dataSheetName}"").Range(tempSheet.UsedRange.Address))
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = True
    If Err <> 0 Then
      MsgBox ""Connection failed. Check service availability""
      Err.Clear
      isRefreshed = False
    Else
      MsgBox ""Connection is refreshed. Use 'Edit Query' to update data""
      isRefreshed = True
    End If
    sheet.Activate
    
End Sub
";
        }

        private static string GetBrowserRefreshDataScript()
        {
            const string makeScriptFileName = "MakeAppleScriptFile.scpt";
            const string makeScriptFunctionName = "CreateAppleScriptFile";
            const string refreshScriptDataFileName = "RefreshData.scpt";
            const string tokenInputId = "token";

            return $@"
#If Mac Then ' Run only on MacOS

    Sub RefreshButtonClick()
        On Error GoTo error_handler
    
         If CheckAppleScriptTaskExcelScriptFile(ScriptFileName:=""{makeScriptFileName}"") = False Then
            messageBox = MsgBox(""Excel is not configured for refreshing data."" & vbCr _
                & vbCr _
                & ""Do you want to download script for setting up Excel?"", vbYesNo, ""There is a problem..."")
            
            If messageBox = vbYes Then
                ActiveWorkbook.FollowHyperlink Address:=""https://google.com""
                MsgBox ""Please run downloaded script, and then try refresh data again""
            End If
                
            Exit Sub
        End If


        Dim AppleScriptTaskFolder As String
        Dim FileName As String
        
        FileName = ""{refreshScriptDataFileName}""
        
        AppleScriptTaskFolder = MacScript(""return POSIX path of (path to desktop folder) as string"")
        AppleScriptTaskFolder = Replace(AppleScriptTaskFolder, ""/Desktop"", """") & ""Library/Application Scripts/com.microsoft.Excel/""
        AppleScriptTaskFolder = AppleScriptTaskFolder & FileName
        
        ' AppleScript for data refreshing:
        ' - open Safari
        ' - navigate to MedtiaTools
        ' - wait user login
        ' - extract token
        Dim ScriptString As String
        ScriptString = ""on GetToken(tokenUrl)"" & Chr(13)
        ScriptString = ScriptString & "" tell application """"Safari"""""" & Chr(13)
        ScriptString = ScriptString & ""     tell window 1"" & Chr(13)
        ScriptString = ScriptString & ""         activate"" & Chr(13)
        ScriptString = ScriptString & ""             set myTab to make new tab"" & Chr(13)
        ScriptString = ScriptString & ""             set URL of myTab to tokenUrl"" & Chr(13)
        ScriptString = ScriptString & ""             set tabIndex to index of myTab"" & Chr(13)
        ScriptString = ScriptString & ""             set current tab to tab tabIndex"" & Chr(13)
        ScriptString = ScriptString & ""             repeat until (URL of myTab contains tokenUrl and source of myTab is not equal to """""""")"" & Chr(13)
        ScriptString = ScriptString & ""                 delay 1"" & Chr(13)
        ScriptString = ScriptString & ""             end repeat"" & Chr(13)
        ScriptString = ScriptString & ""             set tokenTag to do shell script """"awk 'match($0, /<input.*?id=\""""{tokenInputId}\"""".*?value=\"""".*?\""""/){{print substr($0, RSTART,RLENGTH)}}' <<< '"""" & source of myTab & """"'"""""" & Chr(13)
        ScriptString = ScriptString & ""             delete myTab"" & Chr(13)
        ScriptString = ScriptString & ""     end tell"" & Chr(13)
        ScriptString = ScriptString & "" end tell"" & Chr(13)
        ScriptString = ScriptString & "" return tokenTag"" & Chr(13)
        ScriptString = ScriptString & ""end GetToken""
        
        Dim scriptCreationString As String
        scriptCreationString = ScriptString & "";"" & AppleScriptTaskFolder
        
        ' Create script for data refreshing
        RunMyScript = AppleScriptTask(""{makeScriptFileName}"", ""{makeScriptFunctionName}"", scriptCreationString)
        
        ' Get token from apple script
        Dim tokenTag As String
        tokenTag = AppleScriptTask(""{refreshScriptDataFileName}"", ""GetToken"", TokenPageLink)
        
        If Not (tokenTag Like ""*value=""""*"""""") Then
            MsgBox ""Error while data exports""
            Exit Sub
        End If
        
        ' Parse token
        Dim dirtyToken As String
        dirtyToken = Split(tokenTag, ""value="""""")(1)
        Dim Token As String
        Token = Left(dirtyToken, Len(dirtyToken) - 1)
        
        RefreshQueryTable (Token)
        Exit Sub
error_handler:
        MsgBox (""Error while data exports"")
    End Sub
    
    Function CheckAppleScriptTaskExcelScriptFile(ScriptFileName As String) As Boolean
        Dim AppleScriptTaskFolder As String
        Dim TestStr As String
    
        AppleScriptTaskFolder = MacScript(""return POSIX path of (path to desktop folder) as string"")
        AppleScriptTaskFolder = Replace(AppleScriptTaskFolder, ""/Desktop"", """") & _
            ""Library/Application Scripts/com.microsoft.Excel/""
    
        On Error Resume Next
        TestStr = Dir(AppleScriptTaskFolder & ScriptFileName, vbDirectory)
        On Error GoTo 0
        If TestStr = vbNullString Then
            CheckAppleScriptTaskExcelScriptFile = False
        Else
            CheckAppleScriptTaskExcelScriptFile = True
        End If
    End Function
    
#Else ' Run only on Windows
    Sub RefreshButtonClick()
        On Error GoTo error_handler
        Dim Token As String
        
        ' Create Internet Explorer application
        Dim ieObject As Object
        Set ieObject = CreateObject(""InternetExplorer.Application"")
            
        ' Try to navigate to page with token (with redirect to Login page)
        ieObject.Visible = True
        ieObject.Navigate TokenPageLink
        
        ' Hack for local testing (IE opens new instance for intranet locations)
        ' So need to find newly opened instance
        Dim shell As Object
        Dim eachIE As Object
        Do
            Set shell = CreateObject(""Shell.Application"")
            For Each eachIE In shell.Windows
                If InStr(1, eachIE.locationurl, LoginPageLink) Or InStr(1, eachIE.locationurl, TokenPageLink) Then
                    Set ieObject = eachIE
                    Set eachIE = Nothing
                    Set shell = Nothing
                Exit Do
                End If
            Next eachIE
        Loop
            
        ' Do loop while user not redirected to requested page from Login page
        Do Until ieObject.ReadyState = 4 And ieObject.LocationName = ""Excel Data Exports"": DoEvents: Loop
            
        ' Get token from hidden input
        Token = ieObject.Document.getElementById(""{tokenInputId}"").Value
            
        ' Clean up resources
        ieObject.Quit
        Set ieObject = Nothing
    
        RefreshQueryTable (Token)
        
        Exit Sub
error_handler:
        MsgBox (""Error while data exports"")
        
        If Not ieObject Is Nothing Then
            ieObject.Quit
            Set ieObject = Nothing
        End If
        If Not shell Is Nothing Then
            Set shell = Nothing
        End If
        If Not eachIE Is Nothing Then
            Set eachIE = Nothing
        End If
    End Sub
#End If

Sub RefreshQueryTable(Token As String)
    Dim sheet As Worksheet
    Set sheet = Sheets(""Data"")
        
    Dim qt As QueryTable
    Set qt = sheet.QueryTables(""{RefreshDataQueryTableName}"")
        
    qt.Connection = ""URL;"" & ExportUrl & ""&userToken="" & Token
    qt.Refresh
    
    MsgBox ""Data was updated""
    sheet.Activate
End Sub
";
        }

        private string GetPivotScript(string dataDimension)
        {
            return $@"
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets.Add After:= ActiveSheet
    ActiveSheet.Name = ""{ExportViewConstants.UnformattedViewPivotSheetName}""
    Application.DisplayAlerts = True

    Dim pivotSheet As Worksheet
    Set pivotSheet = Sheets(""{ExportViewConstants.UnformattedViewPivotSheetName}"")
    
    Dim dataSheet As Worksheet
    Set dataSheet = Sheets(""{_dataSheetName}"")

    Set PRange = dataSheet.Range(""{dataDimension}"").CurrentRegion

    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:= xlDatabase, SourceData:= PRange)
    Set pTable = pvtCache.CreatePivotTable(TableDestination:= pivotSheet.Cells(3, 1), TableName:= ""PivotTable"")

    'Set pt = pivotSheet.PivotTables(1)
    'For Each pf In pt.ColumnFields
    '    pf.Orientation = xlHidden
    'Next pf
";
        }
    }
}