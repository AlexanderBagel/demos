unit droptarget_main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Menus,

  ShlObj,
  ActiveX,
  FWOleDragDrop;

type
  TDropTargetMainForm = class(TForm)
    Label1: TLabel;
    Memo1: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    FDropTarget: TFWDropTarget;
    procedure OnSetContent(Sender: TObject;
      const FileDescriptor: TFileDescriptor; Data: TStream;
      DragEffect: TFWDragEffects);
  end;

var
  DropTargetMainForm: TDropTargetMainForm;

implementation

{$R *.dfm}

procedure TDropTargetMainForm.FormCreate(Sender: TObject);
begin
  FDropTarget := TFWDropTarget.Create;
  FDropTarget.OnSetContent := OnSetContent;
  FDropTarget.RegisterDragDrop(Self);
end;

procedure TDropTargetMainForm.FormDestroy(Sender: TObject);
begin
  FDropTarget.RevokeDragDrop;
end;

procedure TDropTargetMainForm.OnSetContent(Sender: TObject;
  const FileDescriptor: TFileDescriptor; Data: TStream;
  DragEffect: TFWDragEffects);
var
  FilePath: string;
  Buff: AnsiString;
begin
  FilePath := PChar(@FileDescriptor.cFileName[0]);
  if FileExists(FilePath) then
  begin
    Memo1.Lines.Add('File: "' + FilePath + '"');
    Memo1.Lines.Add('Status: real');
  end
  else
    if DirectoryExists(FilePath) then
    begin
      Memo1.Lines.Add('Folder: "' + FilePath + '"');
      Memo1.Lines.Add('Status: real');
    end
    else
    begin
      Memo1.Lines.Add('File: "' + FilePath + '"');
      Memo1.Lines.Add('Status: virtual');
      SetLength(Buff, Data.Size);
      Data.Read(Buff[1], Data.Size);
      Memo1.Lines.Add('Data:');
      Memo1.Lines.Add(string(Buff));
    end;
end;

end.
