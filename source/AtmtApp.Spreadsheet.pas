unit AtmtApp.SpreadSheet;

{
    Remote control of spreadsheet application
}

interface

{
    AutomateApp4Pas

    Base class for spreadsheet automation
}


uses
    System.UiTypes,
    AtmtApp.OleApp;

type

    /// single cell
    ISpreadCell = interface
        procedure SetFontStyles(AStyle: TFontStyles);
        function PropGetText: string;
        procedure PropSetText(const AVal: string);   
        property Text: string read PropGetText write PropSetText;
    end;

    ISpreadsheetApp = interface(IOleApp)
        //function IsInstalled(): Boolean;  via IOleApp
        procedure OpenDocument(AFileName: string);
        function HasDocument(): Boolean;
        function GetCell(ARow: Integer; ACol: Integer): ISpreadCell; 
        function GetRowCount(): Integer;
        function GetColCount(): Integer;
    end;


    /// Base class for calendars
    TISpreadsheetApp = class abstract(TIOleApp, ISpreadsheetApp)
    protected   //public via interface
        procedure OpenDocument(AFileName: string); virtual; abstract;
        function HasDocument(): Boolean; virtual; abstract;
        function GetCell(ARow: Integer; ACol: Integer): ISpreadCell; virtual; abstract;
        function GetRowCount(): Integer; virtual; abstract;
        function GetColCount(): Integer; virtual; abstract;
    end;

    TISpreadsheetAppClass = class of TISpreadsheetApp;
    
    /// Base functions
    TISpreadCell = class abstract(TInterfacedObject)
    protected   //public via interface
    end;


implementation

uses
    System.SysUtils;



end.
