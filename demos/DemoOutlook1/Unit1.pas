unit Unit1;

interface

uses
    Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
    Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
    AtmtApp.Calendar,
    AtmtApp.CalendarOutlook;


type
    TFormDemo1 = class(TForm)
    GroupBoxOutlook: TGroupBox;
    ButtonTest: TButton;
    ButtonVersion: TButton;
    ButtonCount: TButton;
    ButtonExport: TButton;
    Memo1: TMemo;
    GroupBoxSettings: TGroupBox;
    RadioGroupDriver: TRadioGroup;
    CheckBoxShowApp: TCheckBox;
    procedure ButtonTestClick(Sender: TObject);
    procedure ButtonVersionClick(Sender: TObject);
    procedure ButtonCountClick(Sender: TObject);
    procedure ButtonExportClick(Sender: TObject);
    private
        FApp: ICalendarApp;
        function GetDriverClass(): TICalendarAppClass;
        function ConnectToApp(): ICalendarApp;
    public
    end;

var
  FormDemo1: TFormDemo1;

implementation

uses
    AtmtApp.Utils,
    Mv.LibBase;

{$R *.dfm}


function TFormDemo1.GetDriverClass(): TICalendarAppClass;
begin
    case RadioGroupDriver.ItemIndex of
        0: Result := TIOutlookCalendarApp;
        else
          raise Exception.Create('Unknown driver');
    end;
end;


function TFormDemo1.ConnectToApp(): ICalendarApp;
var
    ShowApp: Boolean;
begin
    if not Assigned(FApp) then
    begin
        FApp := GetDriverClass().Create();
        try
            ShowApp := CheckBoxShowApp.Checked;
            FApp.Connect(ShowApp);
        except
            begin
                FApp := nil;
                raise;
            end;
        end;
    end;
    Result := FApp;
end;


procedure TFormDemo1.ButtonTestClick(Sender: TObject);
var
    App: ICalendarApp;
begin
    App := GetDriverClass().Create();   //check without connect
    ShowMessage(FormatInstalledAndRunning('Outlook', App.IsInstalled(), App.IsRunning()));
end;


procedure TFormDemo1.ButtonVersionClick(Sender: TObject);
begin
    ShowMessage(ConnectToApp().GetVersion());
end;

procedure TFormDemo1.ButtonCountClick(Sender: TObject);
var
    App: ICalendarApp;
    Appointments: ICalAppmtList;
begin
    App := ConnectToApp();
    Appointments := App.GetAppointments();
    ShowMessage(Format('%d appointments', [Appointments.Count]));
end;

procedure TFormDemo1.ButtonExportClick(Sender: TObject);
var
    App: ICalendarApp;
    Appointments: ICalAppmtList;
    Appmt: ICalAppmt;
    I: Integer;
    S: string;
begin
    App := ConnectToApp();
    Appointments := App.GetAppointments();
    Memo1.Lines.Clear;
    Memo1.Lines.Add(Format('Found %d appointments:', [Appointments.Count]));
    Memo1.Lines.Add(Format('----------------------', [Appointments.Count]));

    Memo1.Lines.BeginUpdate;
    try
        for I := 0 to Appointments.Count - 1 do
        begin
            Appmt := Appointments[I];
            S := '* ' + Appmt.ToString();
            S := S + '(';
            S := ECat('Location: ', Appmt.Location);
            S := Cat(S, ECat('ID: ', Appmt.GlobalId), ', ');
            S := S + ')';
            Memo1.Lines.Add(S);
        end;
    finally
        Memo1.Lines.EndUpdate;
    end;

end;



end.
