program droptarget_demo;

uses
  Forms,
  droptarget_main in 'droptarget_main.pas' {DropTargetMainForm},
  FWOleDragDrop in '..\common\FWOleDragDrop.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TDropTargetMainForm, DropTargetMainForm);
  Application.Run;
end.
