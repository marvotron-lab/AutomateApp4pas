unit Unit1;

interface

uses
    Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
    Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
    AtmtApp.Spreadsheet,
    AtmtApp.SpreadsheetHoja;


type
    TFormDemo1 = class(TForm)
    GroupBoxApp: TGroupBox;
    ButtonTest: TButton;
    ButtonVersion: TButton;
    ButtonOpen: TButton;
    ButtonEditCell: TButton;
    Memo1: TMemo;
    GroupBoxSettings: TGroupBox;
    RadioGroupDriver: TRadioGroup;
    CheckBoxShowApp: TCheckBox;
    ButtonGetContent: TButton;
    FileOpenDialog1: TFileOpenDialog;
    LabelDescription: TLabel;
    procedure ButtonTestClick(Sender: TObject);
    procedure ButtonVersionClick(Sender: TObject);
    procedure ButtonOpenClick(Sender: TObject);
    procedure ButtonEditCellClick(Sender: TObject);
    procedure ButtonGetContentClick(Sender: TObject);
    procedure CheckBoxShowAppClick(Sender: TObject);
    private
        FApp: ISpreadsheetApp;
        function GetDriverClass(): TISpreadsheetAppClass;
        function ConnectToApp(): ISpreadsheetApp;
        procedure SetDesc(AText: string);
    public
    end;

var
  FormDemo1: TFormDemo1;

implementation

uses
    AtmtApp.Utils;

{$R *.dfm}

function TFormDemo1.GetDriverClass(): TISpreadsheetAppClass;
begin
    case RadioGroupDriver.ItemIndex of
        0: Result := TIHojaSpreadsheetApp;
        else
          raise Exception.Create('Unknown driver');
    end;
end;

function TFormDemo1.ConnectToApp(): ISpreadsheetApp;
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
        if not (FApp.IsVisible() = ShowApp) then
          FApp.SetVisible(ShowApp);
    end;
    Result := FApp;
end;


procedure TFormDemo1.SetDesc(AText: string);
begin
    LabelDescription.Caption := AText;
end;


procedure TFormDemo1.CheckBoxShowAppClick(Sender: TObject);
begin
    if Assigned(FApp) then
      FApp.SetVisible(CheckBoxShowApp.Checked);
    //else on connect / open
end;


procedure TFormDemo1.ButtonTestClick(Sender: TObject);
var
    App: ISpreadsheetApp;
begin
    SetDesc(ButtonTest.Hint);
    App := GetDriverClass().Create();
    ShowMessage(FormatInstalledAndRunning(GetDriverClass().Classname, App.IsInstalled(), App.IsRunning()));
end;


procedure TFormDemo1.ButtonVersionClick(Sender: TObject);
begin
    SetDesc(ButtonVersion.Hint);
    ShowMessage(ConnectToApp().GetVersion());
end;

procedure TFormDemo1.ButtonOpenClick(Sender: TObject);
var
    App: ISpreadsheetApp;
begin
    SetDesc(ButtonOpen.Hint);
    if FileOpenDialog1.Execute then
    begin
        App := ConnectToApp();
        App.OpenDocument(FileOpenDialog1.FileName);
    end;
end;

procedure TFormDemo1.ButtonEditCellClick(Sender: TObject);
var
    App: ISpreadsheetApp;
    Cell: ISpreadCell;
begin
    SetDesc(ButtonEditCell.Hint);
    App := ConnectToApp();
    //doc must be opened
    Cell := App.GetCell(3,4);
    Cell.Text := Memo1.Text;
end;


procedure TFormDemo1.ButtonGetContentClick(Sender: TObject);
var
    App: ISpreadsheetApp;
    Cell: ISpreadCell;
    Row: Integer;
    Col: Integer;
    S: string;
begin
    SetDesc(ButtonGetContent.Hint);
    Memo1.Lines.Clear;
    App := ConnectToApp();

    Memo1.Lines.BeginUpdate;
    try
        for Row := 0 to App.GetRowCount() do
        begin
            S := '';
            for Col := 0 to App.GetColCount() - 1 do
            begin
                Cell := App.GetCell(Row, Col);
                S := S + Cell.Text + ';'        //this does not consider ; in Cell.Text
            end;
            Memo1.Lines.Add(S);
        end;
    finally
        Memo1.Lines.EndUpdate;
    end;
end;

end.
