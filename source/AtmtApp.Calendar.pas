unit AtmtApp.Calendar;

{
    Remote control of calendar application
}

interface

{
    AutomateApp4Pas

    Base class for calendar automation
}


uses
    AtmtApp.Base,
    AtmtApp.OleApp;

type
    EAtmtAppException = AtmtApp.Base.EAtmtAppException;


    /// Attendee of an appointment or recipient of an email
    IRecipient = interface
        function PropGetName: string;
        function PropGetEmail: string;
        property Name: string read PropGetName;
        property Email: string read PropGetEmail;
    end;

    IRecipientList = interface
        function Count(): Integer;
        function GetItem(AIndex: Integer): IRecipient;
        property Items[Index: Integer]: IRecipient read GetItem; default; // write Put;
    end;

    /// single appointment
    ICalAppmt = interface
        function ToString(): string;
        function GetAttendees(): IRecipientList;
        function PropGetStartTime: TDateTime;
        function PropGetEndTime: TDateTime;
        function PropGetBody: string;
        function PropGetSubject: string;
        function PropGetGlobalId: string;
        function PropGetUniqueGlobalId: string;
        function PropGetLocation: string;
        property StartTime: TDateTime read PropGetStartTime;
        property EndTime: TDateTime read PropGetEndTime;
        property Body: string read PropGetBody;
        property Subject: string read PropGetSubject;
        /// ID unique for this appointment (but not for recurring items) = GlobalAppointmentId in Outlook
        property GlobalId: string read PropGetGlobalId;
        /// Unique ID for every item (includes recurring items) = own format
        property UniqueGlobalId: string read PropGetUniqueGlobalId;
        property Location: string read PropGetLocation;
    end;


    /// list of appointments
    ICalAppmtList = interface
        function Count(): Integer;
        function GetItem(AIndex: Integer): ICalAppmt;
        property Items[Index: Integer]: ICalAppmt read GetItem; default; // write Put;
    end;


    ICalendarApp = interface(IOleApp)
        //function IsInstalled(): Boolean;  via IOleApp
        function GetAppointments(): ICalAppmtList;
        function GetFilteredAppointments(AFrom: TDateTime; ATo: TDateTime;
          AIncludeRecurring: Boolean): ICalAppmtList;
    end;


    /// Base class for calendars
    TICalendarApp = class(TIOleApp, ICalendarApp)
    protected   //public via interface
        function GetAppointments(): ICalAppmtList; virtual; abstract;
        function GetFilteredAppointments(AFrom: TDateTime; ATo: TDateTime;
          AIncludeRecurring: Boolean): ICalAppmtList; virtual; abstract;
    end;

    TICalendarAppClass = class of TICalendarApp;

    /// Base functions
    TICalAppmt= class abstract(TInterfacedObject, ICalAppmt)
    public
        function ToString(): string; override;
    protected   //public via interface
        function GetAttendees(): IRecipientList; virtual; abstract;
        function PropGetStartTime: TDateTime; virtual; abstract;
        function PropGetEndTime: TDateTime; virtual; abstract;
        function PropGetBody: string; virtual; abstract;
        function PropGetSubject: string; virtual; abstract;
        function PropGetGlobalId: string; virtual; abstract;
        function PropGetUniqueGlobalId: string; virtual; abstract;
        function PropGetLocation: string; virtual; abstract;
    end;

const
    RECURRING_YES = True;
    RECURRING_NO = False;
    DESCENDING_YES = True;
    DESCENDING_NO = False;


implementation

uses
    System.SysUtils,
    System.DateUtils;

{ TICalAppmt }

function TICalAppmt.ToString: string;
var
    Dat: TDateTime;
begin
    Dat := DateOf(PropGetStartTime());
    if SameDate(Dat, DateOf(PropGetEndTime)) then
      Result := DateToStr(Dat) + ': ' + TimeToStr(PropGetStartTime()) + ' - ' + TimeToStr(PropGetEndTime())
    else
      Result := DateTimeToStr(PropGetStartTime()) + ' - ' + DateTimeToStr(PropGetEndTime());

    Result := Result + ': ' + PropGetSubject() + ': ' + PropGetBody();
end;

end.
