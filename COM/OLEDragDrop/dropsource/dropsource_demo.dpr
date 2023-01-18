program dropsource_demo;

{$R 'resources.res' 'resources.rc'}

uses
  Forms,
  dropsource_main in 'dropsource_main.pas' {DropSourceMainForm},
  FWOleDragDrop in '..\common\FWOleDragDrop.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TDropSourceMainForm, DropSourceMainForm);
  Application.Run;
end.
