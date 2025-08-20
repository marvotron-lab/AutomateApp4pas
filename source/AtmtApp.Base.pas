unit AtmtApp.Base;

{
    AutomateApp4pas

    Base classes
}

interface

uses
    System.SysUtils;

type
    EAtmtAppException = class(Exception)
    public
        const UNKNOWN = $00000000;
        const NOT_AVAILABLE = $71C30001;        //eg not installed
        const CANNOT_CONNECT = $71C30002;       //for instance RPC server not available
        const NO_DOCUMENT = $71C30003;          //No document available
    private
        FOleClassName: string;
        FErrorCode: Integer;
    protected
        function GetErrMsgForCode(OurErrorCode: Integer): string;
    public
        constructor Create(AOurErrorCode: Integer; AOleClassName: string; AErrorMessageAddon: string = '');
    public
        class procedure RaiseForKnownExceptions(E: Exception; AOleClassName: string;
          AErrorCodeForUnknownExceptions: Integer = UNKNOWN);
        property ErrorCode: Integer read FErrorCode;
    end;

const
    SHOW_YES = True;
    SHOW_NO = False;
    FORCE_YES = True;
    FORCE_NO = False;

implementation

uses
    System.Variants,
    System.Win.ComObj;


{! The exception is defined by AOurErrorCode @see GetErrMsgForCode
   @param AErrorMessagePrefix is added to the error message if not empty
------------------------------------------------------------------------------------------------------------------}
constructor EAtmtAppException.Create(AOurErrorCode: Integer; AOleClassName: string; AErrorMessageAddon: string = '');
var
    ErrMsg: string;
begin
    FOleClassName := AOleClassName;
    FErrorCode := AOurErrorCode;
    if (AOurErrorCode = EAtmtAppException.UNKNOWN) and not AErrorMessageAddon.IsEmpty then
      ErrMsg := AErrorMessageAddon
    else begin
        ErrMsg := GetErrMsgForCode(AOurErrorCode);
        if not AErrorMessageAddon.IsEmpty then
          ErrMsg := ErrMsg + #13#10 + AErrorMessageAddon;
    end;
    inherited Create(ErrMsg);
end;

{! Translate error code
------------------------------------------------------------------------------------------------------------------}
function EAtmtAppException.GetErrMsgForCode(OurErrorCode: Integer): string;
begin
    case OurErrorCode of
        EAtmtAppException.NOT_AVAILABLE:
            Result := Format('Automatable application is not installed on this computer (%s).', [FOleClassName]);
        EAtmtAppException.CANNOT_CONNECT:
            Result := Format('Cannot connect with application (%s).', [FOleClassName]);
        EAtmtAppException.NO_DOCUMENT:
            Result := 'Operation is not possible. The document is not available anymore.';
        else
            begin
                Result := 'An unidentified error occurred: ' + Self.Classname + ': ' + FOleClassName;
                Assert(OurErrorCode = EAtmtAppException.UNKNOWN);
            end;
    end;
end;


{ This can be used in an exception handler to catch known exceptions
  @param AErrorCodeForUnknownExceptions forces to throw an exception for other unknown exceptions
------------------------------------------------------------------------------------------------------------------}
class procedure EAtmtAppException.RaiseForKnownExceptions(E: Exception; AOleClassName: string;
  AErrorCodeForUnknownExceptions: Integer = EAtmtAppException.UNKNOWN);
const
    ERR_INVALID_CLASS_STRING: HRESULT = HRESULT($800401F3);
begin
    if E is EOleSysError then
    begin
        if (E as EOleSysError).ErrorCode = ERR_INVALID_CLASS_STRING then
          //our error message is ignored in TApplication.ShowException with this construct:
          //Exception.RaiseOuterException(EAtmtAppException.Create(EAtmtAppException.NOT_AVAILABLE, AOleClassName));
          raise EAtmtAppException.Create(EAtmtAppException.NOT_AVAILABLE, AOleClassName);
    end
    //Throw error for unknown exceptions if AErrorCodeForUnknownExceptions is specified
    else if not (E is EAtmtAppException) and (AErrorCodeForUnknownExceptions <> EAtmtAppException.UNKNOWN) then
      raise EAtmtAppException.Create(AErrorCodeForUnknownExceptions, AOleClassName, E.Message + '(' + E.ClassName + ')');
end;



end.
