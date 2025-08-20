unit AtmtApp.CalendarOutlook;

{
    AutomateApp4pas for Outlook
}

interface

uses
    AtmtApp.OleApp,
    AtmtApp.Calendar;

type
    //only need to include one unit with these definitions
    IRecipient = AtmtApp.Calendar.IRecipient;
    IRecipientList = AtmtApp.Calendar.IRecipientList;
    ICalendarApp = AtmtApp.Calendar.ICalendarApp;
    ICalAppmtList = AtmtApp.Calendar.ICalAppmtList;
    ICalAppmt = AtmtApp.Calendar.ICalAppmt;
    TConnectResult = AtmtApp.OleApp.TConnectResult;


    TIOutlookCalendarApp = class(TICalendarApp, ICalendarApp)
    protected
        class function GetOleClassName(): string; override;
    protected   //public via interface
        //IOleApp
        function GetVersion(): string; override;
        function IsVisible(): Boolean; override;
        procedure SetVisible(AVisible: Boolean); override;
        //ICalendarApp
        function GetAppointments(): ICalAppmtList; override;
        function GetFilteredAppointments(AFrom: TDateTime; ATo: TDateTime;
          AIncludeRecurring: Boolean): ICalAppmtList; override;
    end;

implementation

uses
    System.SysUtils,
    System.Classes,
    System.Variants,
    System.DateUtils,
    Soap.EncdDecd,
    Mv.LibBase;         //HexToBytes


type
    /// Outlook implementation for ICalAppmtList
    TIOutlookCalAppmtList = class(TInterfacedObject, ICalAppmtList)
    private
        FItems: OleVariant;
    protected   //public via interface
        function Count(): Integer;
        function GetItem(AIndex: Integer): ICalAppmt;
    public
        constructor Create(AOutlookCalendar: OleVariant; AFrom, ATo: TDateTime;
          AIncludeRecurring: Boolean);
    end;

    /// Outlook implementation for ICalAppmt
    TIOutlookCalAppmt = class(TICalAppmt, ICalAppmt)
    private
        FAppmtItem: OleVariant;
    protected   //public via interface
        function GetAttendees(): IRecipientList; override;
        function PropGetStartTime: TDateTime; override;
        function PropGetEndTime: TDateTime; override;
        function PropGetBody: string; override;
        function PropGetSubject: string; override;
        function PropGetGlobalId: string; override;
        function PropGetUniqueGlobalId: string; override;
        function PropGetLocation: string; override;
        class function CalcGlobalUniqueId(const AGlobalAppointmentId: string;
          const AStartDate: TDateTime; const AIsRecurringAppointment: Boolean): string;
    public
        constructor Create(AOutlookAppmtItem: OleVariant);
    end;

    /// Outlook implementation for IRecipientList
    TIOutlookRecipientList = class(TInterfacedObject, IRecipientList)
    private
        FOutlookRecipients: OleVariant;
    protected   //public via interface
        function Count(): Integer;
        function GetItem(AIndex: Integer): IRecipient;
    public
        constructor Create(AOutlookRecipients: OleVariant);
    end;

    /// Outlook implementation for IRecipient
    TIOutlookRecipient = class(TInterfacedObject, IRecipient)
    private
        FRecipient: OleVariant;
    protected   //public via interface
        function PropGetName: string;
        function PropGetEmail: string;
    public
        constructor Create(AOutlookRecipient: OleVariant);
    end;


const
    OUTLOOK_OLE_CLASSNAME = 'Outlook.Application';
    //MSWORD_OLE_CLASSNAME = 'Word.Application';

    olFolderCalendar = 9;

{*****************************************************************************************************************}
{$region 'TIOutlookCalAppmtList'}
{*****************************************************************************************************************}

/// see https://learn.microsoft.com/en-us/office/vba/api/outlook.folder
/// @param AOutlookCalendar must be of Outlook type Folder,
/// @param AFrom only consider start dates >= AFrom. AFrom must be >= Date(0)
/// @param ATo only consider start dates <= ATo. ATo must be >= AFrom
constructor TIOutlookCalAppmtList.Create(AOutlookCalendar: OleVariant; AFrom, ATo: TDateTime;
  AIncludeRecurring: Boolean);
begin
    Assert(not VarIsNull(AOutlookCalendar));
    Assert((AFrom >= 0) or ((AFrom = -1) and (ATo = -1)));
    Assert(AFrom <= ATo);
    //---
    FItems := AOutlookCalendar.Items;
    //From docu: https://learn.microsoft.com/en-us/office/vba/api/outlook.items.includerecurrences
    //you need to sort and filter on appointment items that contain recurring appointments,
    //you must do so in this order: sort the items in ascending order, set IncludeRecurrences to True,
    //and then filter the items.
    FItems.Sort('[Start]'); //ascending
    FItems.IncludeRecurrences := AIncludeRecurring;

    //internal: consider -1  = no value
    if (AFrom <> - 1) and (ATo <> - 1) then
    begin
        FItems.Restrict(Format('[Start] >= "%s" AND [Start] < "%s"',
          [FormatDateTime('dd/mm/yyyy', AFrom), FormatDateTime('dd/mm/yyyy', ATo)]));
    end;
end;

function TIOutlookCalAppmtList.Count: Integer;
begin
    //https://learn.microsoft.com/en-us/office/vba/api/outlook.folder.items
    Result := FItems.Count;
end;

function TIOutlookCalAppmtList.GetItem(AIndex: Integer): ICalAppmt;
begin
    //from docu (https://learn.microsoft.com/en-us/office/vba/api/outlook.folder.items):
    //> The index for the Items collection starts at 1, and the items in the Items collection object are not
    //> guaranteed to be in any particular order.
    //returns https://learn.microsoft.com/en-us/office/vba/api/outlook.items

    //FOutlookCalendar.Items can theoretically be any of
    //
    //AppointmentItem Object
    //ContactItem Object
    //DistListItem Object
    //DocumentItem Object
    //JournalItem Object
    //MailItem Object
    //MeetingItem Object
    //NoteItem Object
    //PostItem Object
    //RemoteItem Object
    //ReportItem Object
    //SharingItem Object
    //StorageItem Object
    //TaskItem Object
    //TaskRequestAcceptItem Object
    //TaskRequestDeclineItem Object
    //TaskRequestItem Object
    //TaskRequestUpdateItem Object
    //
    //but we are expecting an AppointmentItem (because of Namespace.GetDefaultFolder(olFolderCalendar);)
    //not sure if check is necessary
    //https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem
    Result := TIOutlookCalAppmt.Create(FItems.Item[AIndex + 1]);      //our AIndex is zero based, outlook is 1-based
end;

{$endregion}


{*****************************************************************************************************************}
{$region 'TIOutlookCalAppmt'}
{*****************************************************************************************************************}

{ According to https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem
  the following properties are available:
    Actions
    AllDayEvent
    Application
    Attachments
    AutoResolvedWinner
    BillingInformation
    Body
    BusyStatus
    Categories
    Class
    Companies
    Conflicts
    ConversationID
    ConversationIndex
    ConversationTopic
    CreationTime
    DownloadState
    Duration
    End
    EndInEndTimeZone
    EndTimeZone
    EndUTC
    EntryID
    ForceUpdateToAllAttendees
    FormDescription
    GetInspector
    GlobalAppointmentID
    Importance
    InternetCodepage
    IsConflict
    IsRecurring
    ItemProperties
    LastModificationTime
    Location
    MarkForDownload
    MeetingStatus
    MeetingWorkspaceURL
    MessageClass
    Mileage
    NoAging
    OptionalAttendees
    Organizer
    OutlookInternalVersion
    OutlookVersion
    Parent
    PropertyAccessor
    Recipients
    RecurrenceState
    ReminderMinutesBeforeStart
    ReminderOverrideDefault
    ReminderPlaySound
    ReminderSet
    ReminderSoundFile
    ReplyTime
    RequiredAttendees
    Resources
    ResponseRequested
    ResponseStatus
    RTFBody
    Saved
    SendUsingAccount
    Sensitivity
    Session
    Size
    Start
    StartInStartTimeZone
    StartTimeZone
    StartUTC
    Subject
    UnRead
    UserProperties
}


/// @param AOutlookAppmtItem must be of Outlook type AppointmentItem,
///   https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem
constructor TIOutlookCalAppmt.Create(AOutlookAppmtItem: OleVariant);
begin
    FAppmtItem := AOutlookAppmtItem;
end;

function TIOutlookCalAppmt.PropGetBody: string;
begin
    Result := FAppmtItem.Body;
end;

function TIOutlookCalAppmt.PropGetSubject: string;
begin
    Result := FAppmtItem.Subject;
end;

function TIOutlookCalAppmt.PropGetStartTime: TDateTime;
begin
    Result := FAppmtItem.Start;
end;

function TIOutlookCalAppmt.PropGetEndTime: TDateTime;
begin
    Result := FAppmtItem.End;
end;

function TIOutlookCalAppmt.PropGetLocation: string;
begin
    Result := FAppmtItem.Location;
end;

/// List of Attendees of the appointment
/// see [Microsofts description](https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.requiredattendees):
/// "The attendee list should be set by using the Recipients collection."
function TIOutlookCalAppmt.GetAttendees: IRecipientList;
begin
    Result := TIOutlookRecipientList.Create(FAppmtItem.Recipients);
end;

{ @returns the OutlookItem.GlobalAppointmentId
    Multiple recurring items share the GlobalAppointmentId of the master item
------------------------------------------------------------------------------------------------------------------}
function TIOutlookCalAppmt.PropGetGlobalId: string;
begin
    //https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.globalappointmentid
    //read only
    //> each Outlook appointment item is assigned a Global Object ID, a unique global identifier which does not
    //> change [...]
    Result := FAppmtItem.GlobalAppointmentID;

    //Here is information, how generation of a GlobalAppointmentID can be done, given a GUID:
    //https://anvil-of-time.com/programming/vba-processing-ics-calendar-file-attachments/
end;

{ @returns the base64 encoded GlobalAppointmentId,
  or if it is a recurring item TimeStamp + '.' + Base64(GlobalAppointmentId).
  This way uniqueness is ensured.
------------------------------------------------------------------------------------------------------------------}
function TIOutlookCalAppmt.PropGetUniqueGlobalId: string;
begin
    Result := TIOutlookCalAppmt.CalcGlobalUniqueId(PropGetGlobalId, PropGetStartTime, FAppmtItem.IsRecurring);
end;

/// returns a string, that should not exceed 110 chars
class function TIOutlookCalAppmt.CalcGlobalUniqueId(const AGlobalAppointmentId: string;
  const AStartDate: TDateTime; const AIsRecurringAppointment: Boolean): string;
var
  GlobalIdBytes: TBytes;
  EncodedId: string;
  ByteStream: TBytesStream;
  FormattedDate: string;
begin
    // Convert the hex string to a byte array
    GlobalIdBytes := HexToBytes(AGlobalAppointmentId);
    // Encode the byte array in Base64
    ByteStream := TBytesStream.Create(GlobalIdBytes);
    try
        //the conversion AnsiString --> string is save (cause Base64 only uses`'save' chars)
        EncodedId := string(EncodeBase64(ByteStream, ByteStream.Size));
    finally
        ByteStream.Free;
    end;

    if AIsRecurringAppointment then
    begin
        //all recurring Items share the same Id - Add StartDate
        // Format the date as 'yyyymmddhhnnss'
        FormattedDate := FormatDateTime('yyyymmddhhnnss', AStartDate);
        // Combine date and base64-encoded GlobalAppointmentID
        // This should be at maximum 15 + 92 Bytes
        Result := FormattedDate + '.' + EncodedId;
    end
    else //if the Result contains a dot ('.') it is a recurring item
      Result := EncodedId;   //max. 92 Bytes
    //---
    Assert(not AIsRecurringAppointment or Result.Contains('.'));
end;

{$endregion}


{*****************************************************************************************************************}
{$region 'TIOutlookRecipientList'}
{*****************************************************************************************************************}

constructor TIOutlookRecipientList.Create(AOutlookRecipients: OleVariant);
begin
    Assert(not VarIsNull(AOutlookRecipients));
    //---
    FOutlookRecipients := AOutlookRecipients;
end;

function TIOutlookRecipientList.Count: Integer;
begin
    //https://learn.microsoft.com/en-us/office/vba/api/outlook.recipients
    Result := FOutlookRecipients.Count;
end;

function TIOutlookRecipientList.GetItem(AIndex: Integer): IRecipient;
begin
    Result := TIOutlookRecipient.Create(FOutlookRecipients.Item[AIndex + 1]);      //our AIndex is zero based, outlook is 1-based
end;

{$endregion}


{*****************************************************************************************************************}
{$region 'TIOutlookRecipient'}
{*****************************************************************************************************************}

constructor TIOutlookRecipient.Create(AOutlookRecipient: OleVariant);
begin
    FRecipient := AOutlookRecipient;
end;

function TIOutlookRecipient.PropGetEmail: string;
begin
    Result := FRecipient.Email;
end;

function TIOutlookRecipient.PropGetName: string;
begin
    Result := FRecipient.Name;
end;

{$endregion}


class function TIOutlookCalendarApp.GetOleClassName: string;
begin
    Result := OUTLOOK_OLE_CLASSNAME;
end;


function TIOutlookCalendarApp.GetVersion: string;
begin
    Result := OleApp.Application.Version;
end;

function TIOutlookCalendarApp.IsVisible: Boolean;
begin
    Result := OleApp.Visible;
end;

procedure TIOutlookCalendarApp.SetVisible(AVisible: Boolean);
begin
    OleApp.Visible := AVisible;
end;

// see https://www.delphipraxis.net/190615-outlook-kalender-auslesen.html
/// For recurring items only the master item is delivered
function TIOutlookCalendarApp.GetAppointments: ICalAppmtList;
begin
    Result := GetFilteredAppointments(-1, -1, RECURRING_NO);
end;


function TIOutlookCalendarApp.GetFilteredAppointments(AFrom: TDateTime; ATo: TDateTime;
  AIncludeRecurring: Boolean): ICalAppmtList;
var
    NameSpace: OleVariant;
    Calendar: OleVariant;
begin
    //https://learn.microsoft.com/de-de/office/vba/api/outlook.namespace
    NameSpace := OleApp.GetNameSpace('MAPI');
    //https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.getdefaultfolder
    //returns https://learn.microsoft.com/en-us/office/vba/api/outlook.folder
    Calendar := Namespace.GetDefaultFolder(olFolderCalendar);
    Result := TIOutlookCalAppmtList.Create(Calendar, AFrom, ATo, AIncludeRecurring);
end;



end.
