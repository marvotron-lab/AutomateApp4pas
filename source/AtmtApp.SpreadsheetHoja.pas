unit AtmtApp.SpreadsheetHoja;

{
    AutomateApp4pas for automation via HojaCalc / Delphi-Spreadsheets unit
}

interface

uses
    AtmtApp.OleApp,
    AtmtApp.Spreadsheet,
    uHojaCalc;


type
    //only need to include one unit with these definitions
    ISpreadsheetApp = AtmtApp.Spreadsheet.ISpreadsheetApp;
    ISpreadCell = AtmtApp.Spreadsheet.ISpreadCell;
    TConnectResult = AtmtApp.OleApp.TConnectResult;

    TIHojaSpreadsheetApp = class(TISpreadsheetApp, ISpreadsheetApp)
    private
        FHojaCalc: THojaCalc;
        FShowApp: Boolean;
    protected   //public via interface
        //IOleApp
        function InternalConnect(AShowApp: Boolean = True): TConnectResult; override;
        function IsInstalled(): Boolean; override;
        function IsRunning(): Boolean; override;
        function GetVersion(): string; override;
        function IsVisible(): Boolean; override;
        procedure SetVisible(AVisible: Boolean); override;
        //ISpreadSheetApp
        procedure OpenDocument(AFileName: string); override;
        procedure CloseDocument();
        function HasDocument(): Boolean; override;
        function GetCell(ARow: Integer; ACol: Integer): ISpreadCell; override;
        function GetRowCount(): Integer; override;
        function GetColCount(): Integer; override;

    public
        destructor Destroy; override;
    end;

implementation

uses
    System.SysUtils,
    System.Variants,
    System.UiTypes,
    AtmtApp.Utils;

type
    /// implementation for ISpreadCell via HojaCalc unit
    TIHojaSpreadCell = class(TInterfacedObject, ISpreadCell)
    private
        FHoja: ISpreadsheetApp;
        FHojaCalc: THojaCalc;
        FRow: Integer;
        FCol: Integer;
        function GetReference(): THojaCalc;
    protected   //public via interface
        procedure SetFontStyles(AStyles: TFontStyles);
        function PropGetText: string;
        procedure PropSetText(const AVal: string);
    public
        constructor Create(AHojaSpreadsheet: ISpreadsheetApp; ARow: Integer; ACol: Integer);
    end;

{ TIHojaSpreadCell }

constructor TIHojaSpreadCell.Create(AHojaSpreadsheet: ISpreadsheetApp; ARow: Integer; ACol: Integer);
begin
    //hold reference to not invalidate pointer to THojaCalc
    FHoja := AHojaSpreadsheet;
    //keep reference and throw error if not available
    FHojaCalc := GetReference();
    //HojaCalc-Rows are 1-based, we are zero based
    FRow := ARow + 1;
    FCol := ACol + 1;
end;


function TIHojaSpreadCell.GetReference: THojaCalc;
begin
    //access private variable
    Result := (FHoja as TIHojaSpreadsheetApp).FHojaCalc;
    if not Assigned(Result) or ((Result <> FHojaCalc) and Assigned(FHojaCalc)) then
      //this might happen if the document has been closed programmatically
      raise EAtmtAppException.Create(EAtmtAppException.NO_DOCUMENT, 'hojacalc');
end;

function TIHojaSpreadCell.PropGetText: string;
begin
    Result := GetReference().CellText[FRow, FCol];
end;

procedure TIHojaSpreadCell.PropSetText(const AVal: string);
begin
    GetReference().CellText[FRow, FCol] := AVal;
end;

/// Hoja does not allow to reset font style (eg from bold to normal)
procedure TIHojaSpreadCell.SetFontStyles(AStyles: TFontStyles);
var
    Hoja: THojaCalc;
begin
    Hoja := GetReference();
    if TFontStyle.fsBold in AStyles then
      Hoja.Bold(FRow, FCol);
    if TFontStyle.fsItalic in AStyles then
      Hoja.Italic(FRow, FCol);
    if TFontStyle.fsUnderline in AStyles then
      Hoja.Underline(FRow, FCol, ulSingle);
end;

{ TIHojaSpreadsheetApp }

destructor TIHojaSpreadsheetApp.Destroy;
begin
    if Assigned(FHojaCalc) then
    begin
        if FShowApp then
        begin
            //Document has been opened visually
            //Document has not been closed
            //so keep the document opened
            FHojaCalc.KeepAlive := True;
        end
        else
          FHojaCalc.KeepAlive := False;
    end;

end;

function TIHojaSpreadsheetApp.InternalConnect(AShowApp: Boolean = True): TConnectResult;
begin
    //do nothing on connect / only on Open / New
    FShowApp := AShowApp;
    //There will always be an new instance with HojaCalc
    Result := TConnectResult.NewInstance;
end;

function TIHojaSpreadsheetApp.IsInstalled: Boolean;
begin
    //This is not directly available
    Result := True;
end;

function TIHojaSpreadsheetApp.IsRunning: Boolean;
begin
    if Assigned(FHojaCalc) then
      Result := FHojaCalc.StillConnectedToApp
    else
      Result := False;
end;

function TIHojaSpreadsheetApp.GetVersion: string;
begin
    if Assigned(FHojaCalc) then
      Result := FHojaCalc.Programa.Application.Version
    else
      Result := inherited;
end;

function TIHojaSpreadsheetApp.IsVisible: Boolean;
begin
    if Assigned(FHojaCalc) then
      Result := FHojaCalc.Visible
    else
      Result := False;
end;

procedure TIHojaSpreadsheetApp.SetVisible(AVisible: Boolean);
begin
    if Assigned(FHojaCalc) then
      FHojaCalc.Visible := AVisible;
    FShowApp := AVisible;
end;


procedure TIHojaSpreadsheetApp.OpenDocument(AFileName: string);
begin
    FHojaCalc := THojaCalc.Create(AFileName, FShowApp, False);   //False = do not reuse existing instance
    //There is a bug in THojaCalc.LoadDoc:
    //m_vActiveSheet := ActivateSheetByIndex(1);    //this is Boolean
    //WORKAROUND: set m_vActiveSheet again:
    FHojaCalc.ActivateSheetByIndex(1);
end;

procedure TIHojaSpreadsheetApp.CloseDocument;
begin
    if Assigned(FHojaCalc) then
    begin
        FHojaCalc.KeepAlive := False;
        FreeAndNil(FHojaCalc);
    end;
end;

function TIHojaSpreadsheetApp.HasDocument: Boolean;
begin
    //in HojaCalc a document is always opened
    Result := Assigned(FHojaCalc);
end;

function TIHojaSpreadsheetApp.GetCell(ARow: Integer; ACol: Integer): ISpreadCell;
begin
    Result := TIHojaSpreadCell.Create(Self, ARow, ACol);
end;

function TIHojaSpreadsheetApp.GetRowCount: Integer;
begin
    if Assigned(FHojaCalc) then
      //HojaCalc is 1-based
      Result := FHojaCalc.LastRow
    else
      Result := 0;
end;

function TIHojaSpreadsheetApp.GetColCount: Integer;
begin
    if Assigned(FHojaCalc) then
      //HojaCalc is 1-based
      Result := FHojaCalc.LastCol
    else
      Result := 0;
end;


end.
