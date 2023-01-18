unit dropsource_main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ShlObj, ActiveX, Menus, CheckLst, Vcl.ExtCtrls,
  PngImage,
  FWOleDragDrop;

type
  TDropSourceMainForm = class(TForm)
    Label1: TLabel;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    CheckListBox1: TCheckListBox;
    Label2: TLabel;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure CheckListBox1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Image1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    FPresentImagePath: string;
    // Флаг, содержащий в себе статус опереции перетаскивания (идет или нет)
    FDropStarted: Boolean;
    // обработчик события для получения виртуальных данных
    procedure OnGetContent(Sender: TObject;
      const FileDescriptor: TFileDescriptor; Data: TStream);
    // обработчик события завершения операции перетаскивания
    procedure OnDropEnd(Sender: TObject);
  end;

var
  DropSourceMainForm: TDropSourceMainForm;

implementation

{$R *.dfm}

const
  Title = 'FWOleDragDrop Source demo';

//
//  Включаем чеки на элементах, которые мы можем перетащить или скопировать
// =============================================================================
procedure TDropSourceMainForm.FormCreate(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to CheckListBox1.Count - 1 do
    CheckListBox1.Checked[I] := True;
  FPresentImagePath := ExtractFilePath(ParamStr(0)) + 'image.png';
  Image1.Picture.LoadFromFile(FPresentImagePath);
end;

procedure TDropSourceMainForm.Image1MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  DropSource: TFWDragDropSourceProvider;
begin
  if Button <> mbLeft then Exit;
  Caption := Title + ' [In Drag Loop]';
  // инициализируем флаг операции перетаскивания
  FDropStarted := True;
  // создаем провайдер, обеспечивающий операцию перетаскивания
  DropSource := TFWDragDropSourceProvider.Create;
  // выставляем флаг - отображать диалог копирования
  DropSource.ShowUIDialog := True;

  // обработчик для отдачи контента файла не назначаем, т.к. файл присутствует
  // на диске и его содержимое будет получаться штатным способом
  // посредством дескриптора CF_HDROP

  // назначаем обработчик завершения операции
  DropSource.OnDropEnd := OnDropEnd;
  // уведомляем провайдера о файлах, которые мы должны выгрузить
  DropSource.AddFile(FPresentImagePath);
  // и инициализируем операцию перетаскивания
  DropSource.ExecuteDragDrop;
end;

//
//  Инициализация процесса перетаскивания
// =============================================================================
procedure TDropSourceMainForm.CheckListBox1MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  I: Integer;
  DropSource: TFWDragDropSourceProvider;
begin
  if Button <> mbLeft then Exit;
  Caption := Title + ' [In Drag Loop]';
  // инициализируем флаг операции перетаскивания
  FDropStarted := True;
  // создаем провайдер, обеспечивающий операцию перетаскивания
  DropSource := TFWDragDropSourceProvider.Create;
  // выставляем флаг - отображать диалог копирования
  DropSource.ShowUIDialog := True;
  // назначаем обработчик, в котором будем отдавать данные виртуального файла (из ресурсов)
  DropSource.OnGetContent := OnGetContent;
  // назначаем обработчик завершения операции
  DropSource.OnDropEnd := OnDropEnd;
  // уведомляем провайдера о файлах, которые мы должны выгрузить
  for I := 0 to CheckListBox1.Count - 1 do
    if CheckListBox1.Checked[I] then
      DropSource.AddFile(CheckListBox1.Items[I]);
  // и инициализируем операцию перетаскивания
  DropSource.ExecuteDragDrop;
end;

//
//  Финализация процесса перетаскивания (по любому событию, успешному или нет)
// =============================================================================
procedure TDropSourceMainForm.OnDropEnd(Sender: TObject);
begin
  FDropStarted := False;
  Caption := Title;
end;

//
//  Копирование данных о виртуальных файлах в буфер обмена
// =============================================================================
procedure TDropSourceMainForm.N1Click(Sender: TObject);
var
  I: Integer;
  CopySource: TFWCopyDataProvider;
begin
  // создаем провайдер, обеспечивающий операцию перетаскивания
  CopySource := TFWCopyDataProvider.Create;
  // выставляем флаг - отображать диалог копирования
  CopySource.ShowUIDialog := True;
  // назначаем обработчик, в котором будем отдавать данные виртуального файла (из ресурсов)
  CopySource.OnGetContent := OnGetContent;
  // уведомляем провайдера о файлах, которые мы должны выгрузить
  for I := 0 to CheckListBox1.Count - 1 do
    if CheckListBox1.Checked[I] then
      CopySource.AddFile(CheckListBox1.Items[I]);
  // и копируем все в буфер обмена, откуда внешнее приложение сможет все забрать
  CopySource.CopyToClibboard;
end;

//
//  Основной обработчик события получения данных виртуальных файлов
// =============================================================================
procedure TDropSourceMainForm.OnGetContent(Sender: TObject;
  const FileDescriptor: TFileDescriptor; Data: TStream);
var
  I: Integer;
  R: TResourceStream;
begin
  // кто-то снаружи запросил содержимое виртуального файла
  // отдаем его, ориентируясь на его имя
  for I := 0 to 4 do
    if PChar(@FileDescriptor.cFileName[0]) = CheckListBox1.Items[I] then
    begin
      R := TResourceStream.Create(HInstance, 'RES' + IntToStr(I + 1),
        'DRAG_RES');
      try
        Data.CopyFrom(R, 0);
      finally
        R.Free;
      end;
    end;
end;

end.
