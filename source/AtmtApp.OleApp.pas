unit AtmtApp.OleApp;

{
    AutomateApp4Pas

    Base class for Ole Automation Apps
}


interface

uses
    System.SysUtils,
    AtmtApp.Base;


type
    EAtmtAppException = AtmtApp.Base.EAtmtAppException;


    {$SCOPEDENUMS ON}
    TConnectResult = (
        NewInstance,
        WasRunning
    );
    {$SCOPEDENUMS OFF}


    IOleApp = interface
        /// Connect to app, throw error if not possible
        /// @return TConnectResult.NewInstance, if a new instance has been opened
        function Connect(AShowApp: Boolean = True): TConnectResult;
        function IsInstalled(): Boolean;
        function IsRunning(): Boolean;
        function GetVersion(): string;
        function IsVisible(): Boolean;
        procedure SetVisible(AVisible: Boolean);
    end;


    TIOleApp = class abstract(TInterfacedObject, IOleApp)
    private
        FOleApp: OleVariant;
        //FWasRunning: Boolean;       //True, if the application has already been running on connect
    protected
        class function GetOleClassName(): string; virtual; abstract;    //eg  'Outlook.Application'
        //Interface functions
        function Connect(AShowApp: Boolean = True): TConnectResult;     //not virtual - override InternalConnect
        function IsInstalled(): Boolean; virtual;
        function IsRunning(): Boolean; virtual;
        function GetVersion(): string; virtual;
        //
        function IsVisible(): Boolean; virtual; abstract;
        procedure SetVisible(AVisible: Boolean); virtual; abstract;
        //
        function InternalConnect(AShowApp: Boolean = True): TConnectResult; virtual;
        function IsInitialized(): Boolean; virtual;
        //
        function PropGetOleApp(): OleVariant;
        property OleApp: OleVariant read PropGetOleApp;
    public
        constructor Create; virtual;
    end;




implementation

uses
    System.Variants,
    AtmtApp.Utils;

{
------------------------------------------------------------------------------------------------------------------}
constructor TIOleApp.Create;
begin
    inherited;
    //no inits atm
end;

{
------------------------------------------------------------------------------------------------------------------}
function TIOleApp.IsInitialized: Boolean;
begin
    Result := not VarIsClear(FOleApp);
end;

{
------------------------------------------------------------------------------------------------------------------}
function TIOleApp.PropGetOleApp(): OleVariant;
begin
    if not IsInitialized() then
    begin
        InternalConnect();
        //InternalConnect must throw error if it fails
        Assert(IsInitialized());
    end;

    Result := FOleApp;
end;

{! connect to ole provider
   @return True, if a new instance has been opened
------------------------------------------------------------------------------------------------------------------}
function TIOleApp.InternalConnect(AShowApp: Boolean = True): TConnectResult;
var
    ResWasRunning: Boolean;
begin
    Assert(not IsInitialized());
    //---
    //Trying to get the App handle
    try
        FOleApp := GetOrCreateOleObject(GetOleClassName(), ResWasRunning);
    except on E: Exception do
        begin
            VarClear(FOleApp);
            Assert(not IsInitialized());
            //try to find known exceptions, throw CANNOT_CONNECT for other
            EAtmtAppException.RaiseForKnownExceptions(E, GetOleClassName(), EAtmtAppException.CANNOT_CONNECT);
            raise;
        end;
    end;

    if AShowApp then
      SetVisible(AShowApp);


    if ResWasRunning then
      Result := TConnectResult.WasRunning
    else
      Result := TConnectResult.NewInstance;
end;


{! Connect to app, throw error if not possible
   @return True, if a new instance has been opened
------------------------------------------------------------------------------------------------------------------}
function TIOleApp.Connect(AShowApp: Boolean = True): TConnectResult;
begin
    if not IsInitialized() then
      Result := InternalConnect(AShowApp)
    else
      Result := TConnectResult.WasRunning;
end;


{! @return True, if installed
------------------------------------------------------------------------------------------------------------------}
function TIOleApp.IsInstalled: Boolean;
begin
    Result := IsOleObjectAvailable(GetOleClassName());
end;

{! @return True, if running
------------------------------------------------------------------------------------------------------------------}
function TIOleApp.IsRunning: Boolean;
begin
    Result := IsOleObjectActive(GetOleClassName());
end;

{! @return Version string
------------------------------------------------------------------------------------------------------------------}
function TIOleApp.GetVersion: string;
begin
    Result := 'Unknown';
end;


end.
