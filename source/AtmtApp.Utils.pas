unit AtmtApp.Utils;

{
    AutomateApp4pas

    Utils for ole automation
}

interface

uses
    AtmtApp.Base,
    System.SysUtils;

type
    EAtmtAppException = AtmtApp.Base.EAtmtAppException;

function IsOleObjectAvailable(const AClassName: string): Boolean;
function IsOleObjectActive(const AClassName: string; AForceAvailable: Boolean = False): Boolean;
function GetOrCreateOleObject(const AClassName: string; out ResultWasRunning: Boolean): OleVariant;

function FormatInstalledAndRunning(AAppName: string; AInstalled: Boolean; ARunning: Boolean): string;


implementation

uses
    System.Variants,
    System.Win.ComObj,
    WinApi.Ole2;        //TCLSID



// see https://www.delphipraxis.net/197623-pruefen-ob-excel-installiert-ist.html
/// True if the application is installed and OLE is available
function IsOleObjectAvailable(const AClassName: string): Boolean;
var
    ClassID: TCLSID;
begin
    Result := Succeeded(CLSIDFromProgID(PWideChar(WideString(AClassName)), ClassID));
end;


// see https://www.delphipraxis.net/197623-pruefen-ob-excel-installiert-ist.html
/// True if the application ist running
function IsOleObjectActive(const AClassName: string; AForceAvailable: Boolean = False): Boolean;
var
    ClassID: TCLSID;
    Unknown: IUnknown;
begin
    Result := False;
    if IsOleObjectAvailable(AClassName) then
      Result := Succeeded(GetActiveObject(ClassID, nil, Unknown))
    else  //handle without catching EOleSysException Exception
      if AForceAvailable then
        raise EAtmtAppException.Create(EAtmtAppException.NOT_AVAILABLE, AClassName);
end;


//see eg https://en.delphipraxis.net/topic/3098-sending-email-and-compose-via-outlook/
/// Returns the running application or creates new instance
function GetOrCreateOleObject(const AClassName: string; out ResultWasRunning: Boolean): OleVariant;
begin
    ResultWasRunning := IsOleObjectActive(AClassName, FORCE_YES);
    try
        if not ResultWasRunning then
          Result := CreateOleObject(AClassName)
        else
          Result := GetActiveOleObject(AClassName);
    except on E: Exception do
        begin
            EAtmtAppException.RaiseForKnownExceptions(E, AClassName);
            raise;
        end;
    end;
end;


///for demos
function FormatInstalledAndRunning(AAppName: string; AInstalled: Boolean; ARunning: Boolean): string;
var
    InstalledStr: string;
    RunningStr: string;
begin
    if AInstalled then
    begin
        InstalledStr := 'available';
        if ARunning then
          RunningStr := 'running'
        else
          RunningStr := 'not runnig';
    end
    else begin
        InstalledStr := 'not available';
        RunningStr := '';
    end;

    Result := Format('%s is %s.', [AAppName, InstalledStr]);
    if RunningStr <> '' then
      Result := Result + ' ' + Format('%s is %s.', [AAppName, RunningStr]);
end;

end.
