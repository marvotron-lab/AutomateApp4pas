program DemoSpreadsheet1;

uses
  Vcl.Forms,
  Unit1 in 'Unit1.pas' {FormDemo1};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFormDemo1, FormDemo1);
  Application.Run;
end.
