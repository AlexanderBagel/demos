////////////////////////////////////////////////////////////////////////////////
//
//  ****************************************************************************
//  * Unit Name : FWOleDragDrop
//  * Purpose   : Набор классов реализующих передачу данных через OLE
//  * Author    : Александр (Rouse_) Багель
//  * Copyright : © Fangorn Wizards Lab 1998 - 2023.
//  * Version   : 1.05
//  * Home Page : http://rouse.drkb.ru
//  * Home Blog : http://alexander-bagel.blogspot.ru
//  * Git       : https://github.com/AlexanderBagel/
//  ****************************************************************************
//

unit FWOleDragDrop;

interface

uses
  Types,
  Windows,
  Classes,
  Controls,
  SysUtils,
  ActiveX,
  ShellAPI,
  ShlObj,
  ComObj;

  {$WARN SYMBOL_PLATFORM OFF}

type
  TFWDragEffect = (deNone, deCopy, deMove, deLink, deScroll);
  TFWDragEffects = set of TFWDragEffect;

  // Немного более удобная структура чем TFileGroupDescriptor
  PFileGroupDescriptorEx = ^TFileGroupDescriptorEx;
  TFileGroupDescriptorEx = record
    cItems: UINT;
    fgd: array of TFileDescriptor;
  end;

  TFWOnGetContentEvent = procedure(Sender: TObject;
    const FileDescriptor: TFileDescriptor; Data: TStream) of object;
  TFWOnCopyResult = procedure(Sender: TObject;
    CopyStyle: TFWDragEffects) of object;
  TFWOnDragOverEvent = procedure(Sender: TObject; pt: TPoint;
    Files: TFileGroupDescriptorEx; DragState: TDragState;
    var DragEffect: TFWDragEffects) of object;
  TFWOnSetContentEvent = procedure(Sender: TObject;
    const FileDescriptor: TFileDescriptor; Data: TStream;
    DragEffect: TFWDragEffects) of object;

  // Интерфейс предоставляет данные о форматах обьектов
  TFWEnumFormatEtc = class(TInterfacedObject, IEnumFormatEtc)
  private
    Fcch: Integer;
    FFormatsEtc: PFormatEtc;
    FPosition: Integer;
  protected
    { IEnumFormatEtc }
    function Next(celt: Longint; out elt;
      pceltFetched: PLongint): HResult; stdcall;
    function Skip(celt: Longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out Enum: IEnumFormatEtc): HResult; stdcall;
  public
    constructor Create(FormatsEtc: PFormatEtc; cch: Integer;
      Position: Integer = 0);
  end;

  // Класс реализующий виртуальный IDataObject применяемый для
  // операций копирования в буффер обмена и для DragDrop
  TFWDataObject = class(TInterfacedObject, IDataObject)
  private
    FContent: TFWOnGetContentEvent;
    FFormatEtc: array of TFormatEtc;
    FValues: TFileGroupDescriptorEx;
    FShowUIDialog: Boolean;
    FSupportedDragEffect: TFWDragEffects;
    FEndCopy: TFWOnCopyResult;
    procedure SetShowUIDialog(const Value: Boolean);
  protected
  { IDataObject }
    function GetData(const formatetcIn: TFormatEtc; out medium: TStgMedium):
      HResult; stdcall;
    function GetDataHere(const formatetc: TFormatEtc; out medium: TStgMedium):
      HResult; stdcall;
    function QueryGetData(const formatetc: TFormatEtc): HResult; stdcall;
    function GetCanonicalFormatEtc(const formatetc: TFormatEtc;
      out formatetcOut: TFormatEtc): HResult; stdcall;
    function SetData(const formatetc: TFormatEtc; var medium: TStgMedium;
      fRelease: BOOL): HResult; stdcall;
    function EnumFormatEtc(dwDirection: Longint; out enumFormatEtc:
      IEnumFormatEtc): HResult; stdcall;
    function DAdvise(const formatetc: TFormatEtc; advf: Longint;
      const advSink: IAdviseSink; out dwConnection: Longint): HResult; stdcall;
    function DUnadvise(dwConnection: Longint): HResult; stdcall;
    function EnumDAdvise(out enumAdvise: IEnumStatData): HResult;
      stdcall;
  protected
    procedure DoEndCopy(CopyStyle: TFWDragEffects); virtual;
    procedure DoGetContent(FileDescriptor: TFileDescriptor; Data: TStream); virtual;
    function IsValidFormat(formatetc: TFormatEtc): HResult; virtual;
  public
    constructor Create;
    procedure AddDescriptor(Value: TFileDescriptor);
    procedure AddFile(const FileName: string); overload;
    procedure AddFile(const FileName: string; FileSize: Int64); overload;
    procedure AddFolder(const FileName: string);
    property ShowUIDialog: Boolean read FShowUIDialog write SetShowUIDialog;
    property OnGetContent: TFWOnGetContentEvent read FContent write FContent;
    property OnEndCopy: TFWOnCopyResult read FEndCopy write FEndCopy;
  end;

  // Класс реализующий виртуальный IDropTarget - приемник при операции DragDrop
  TFWDropTarget = class;

  IFWDropTarget = interface
  ['{5746D886-DF33-4912-8133-2A38DA7B3443}']
    function GetInstance: TFWDropTarget;
  end;

  TFWDropTarget = class(TInterfacedObject, IDropTarget, IFWDropTarget)
  private
    FDropTarget: TWinControl;
    FDragOver: TFWOnDragOverEvent;
    FFiles: TFileGroupDescriptorEx;
    FVirtualFiles: Boolean;
    FContent: TFWOnSetContentEvent;
    FDE: TFWDragEffects;
  protected
    { IDropTarget }
    function DragEnter(const dataObj: IDataObject; grfKeyState: Longint;
      pt: TPoint; var dwEffect: Longint): HResult; stdcall;
    function DragOver(grfKeyState: Longint; pt: TPoint;
      var dwEffect: Longint): HResult; stdcall;
    function DragLeave: HResult; stdcall;
    function Drop(const dataObj: IDataObject; grfKeyState: Longint; pt: TPoint;
      var dwEffect: Longint): HResult; stdcall;
    {IFWDropTarget}
    function GetInstance: TFWDropTarget;
  protected
    procedure ClearFiles;
    procedure DoDragOver(pt: TPoint; Files: TFileGroupDescriptorEx;
      DragState: TDragState; var DragEffect: TFWDragEffects);
    procedure DoSetContent(
      const FileDescriptor: TFileDescriptor; Data: TStream;
      DragEffect: TFWDragEffects);
    procedure ExtractFileNames(const dataObj: IDataObject; pt: TPoint);
  public
    constructor Create;
    procedure RegisterDragDrop(DropTarget: TWinControl);
    procedure RevokeDragDrop;
    property SupportedDragEffect: TFWDragEffects read FDE write FDE;
    property OnDragOver: TFWOnDragOverEvent read FDragOver write FDragOver;
    property OnSetContent: TFWOnSetContentEvent read FContent write FContent;
  end;

  // Класс реализующий IDropSource, стартует операцию перетаскивания
  // при помощи вызова метода ExecuteDragDrop
  TFWDragDropSourceProvider = class(TFWDataObject, IDropSource)
  private
    FDropEnd: TNotifyEvent;
  protected
  { IDropSource }
    function QueryContinueDrag(fEscapePressed: BOOL;
      grfKeyState: Longint): HResult; stdcall;
    function GiveFeedback(dwEffect: Longint): HResult; stdcall;
  protected
    procedure DoDropEnd; virtual;
  public
    function ExecuteDragDrop(
      const DragEffect: TFWDragEffects = [deCopy]): HRESULT;
    property OnDropEnd: TNotifyEvent read FDropEnd write FDropEnd;
  end;

  // Класс - враппер над TFWDataObject, предоставляет методы копирования
  // инициализированного IDataObject в буффер обмена и последующей
  // программной вставки в папку
  TFWCopyDataProvider = class(TFWDataObject)
  public
    function CopyToClibboard: HRESULT;
    function PasteToFolder(const Path: string): HRESULT;
    procedure ClearClipboard;
  end;

  TDropImageType = (
    DROPIMAGE_INVALID = -1,
    DROPIMAGE_NONE = 0,
    DROPIMAGE_COPY = DROPEFFECT_COPY,
    DROPIMAGE_MOVE = DROPEFFECT_MOVE,
    DROPIMAGE_LINK = DROPEFFECT_LINK,
    DROPIMAGE_LABEL = 6,
    DROPIMAGE_WARNING = 7);

  PDropDescription = ^TDropDescription;
  TDropDescription = record
    DROPIMAGETYPE: TDropImageType;
    szMessage: array [0..MAX_PATH - 1] of WideChar;
    szInsert: array [0..MAX_PATH - 1] of WideChar;
  end;

  // Класс работает с буфером обмена и возвращает список файлов
  // реально существующих на диске, которые были скопированы в буфер обмена из проводника
  TFWPasteDataProvider = class
  public
    class function IsFileListAvailable: Boolean;
    class function GetFileListFromClipboard: TStringList;
  end;

implementation

const
  FD_PROGRESSUI = $4000;
  DRAG_EFFECTS: array [TFWDragEffect] of DWORD = (DROPEFFECT_NONE,
    DROPEFFECT_COPY, DROPEFFECT_MOVE, DROPEFFECT_LINK, DROPEFFECT_SCROLL);
  TYMED_ARRAY: array [0..6] of Integer = (TYMED_HGLOBAL, TYMED_FILE,
    TYMED_ISTREAM, TYMED_ISTORAGE, TYMED_GDI, TYMED_MFPICT, TYMED_ENHMF);

var
  CF_FILECONTENTS: Cardinal;
  CF_FILEDESCRIPTOR: Cardinal;
  CF_PERFORMEDDROPEFFECT: Cardinal;
  CF_PREFERREDDROPEFFECT: Cardinal;
  CF_PASTESUCCEEDED: Cardinal;

function GetFormatEtc(cfFormat: TClipFormat; ptd: PDVTargetDevice;
  dwAspect, lindex, tymed: Longint): TFormatEtc; overload;
begin
  Result.cfFormat := cfFormat;
  Result.ptd := ptd;
  Result.dwAspect := dwAspect;
  Result.lindex := lindex;
  Result.tymed := tymed;
end;

function GetFormatEtc(cfFormat: TClipFormat): TFormatEtc; overload;
begin
  Result := GetFormatEtc(cfFormat, nil,
    DVASPECT_CONTENT, -1, TYMED_HGLOBAL);
end;

function DWORDToEffect(Value: DWORD;
  SupportedDragEffect: TFWDragEffects): TFWDragEffects;
var
  iDragEffect: TFWDragEffect;
begin
  Result := [];
  for iDragEffect := Low(TFWDragEffect) to High(TFWDragEffect) do
    if (Value and DRAG_EFFECTS[iDragEffect]) = DRAG_EFFECTS[iDragEffect] then
      if iDragEffect in SupportedDragEffect then
        Include(Result, iDragEffect);
  Exclude(Result, deNone);
end;

function EffectToDWORD(Value: TFWDragEffects): DWORD;
var
  iDragEffect: TFWDragEffect;
begin
  Result := DROPEFFECT_NONE;
  for iDragEffect := Low(TFWDragEffect) to High(TFWDragEffect) do
    if iDragEffect in Value then
      Result := Result or DRAG_EFFECTS[iDragEffect];
end;

{ TFWEnumFormatEtc }

//  Функция создает копию самого себя
// =============================================================================
function TFWEnumFormatEtc.Clone(out Enum: IEnumFormatEtc): HResult;
begin
  Enum := TFWEnumFormatEtc.Create(FFormatsEtc, Fcch, FPosition);
  Result := S_OK;
end;

//  Первоначальная инициализация класса
// =============================================================================
constructor TFWEnumFormatEtc.Create(FormatsEtc: PFormatEtc; cch,
  Position: Integer);
begin
  inherited Create;
  FFormatsEtc := FormatsEtc;
  Fcch := cch;
  FPosition := Position;
end;

//  Функция копирует celt элементов в буффер elt,
//  начиная с текущей позиции FPosition
// =============================================================================
function TFWEnumFormatEtc.Next(celt: Integer; out elt;
  pceltFetched: PLongint): HResult;
var
  I: Integer;
  CurrentCursor, EltCursor: PFormatEtc;
begin
  if celt + FPosition > Fcch then
  begin
    Result := S_FALSE;
    Exit;
  end;
  EltCursor := PFormatEtc(@elt);
  if IsBadWritePtr(EltCursor, celt * SizeOf(TFormatEtc)) then
  begin
    Result := S_FALSE;
    Exit;
  end;
  CurrentCursor := FFormatsEtc;
  Inc(CurrentCursor, FPosition);
  for I := 0 to celt - 1 do
  begin
    EltCursor^ := CurrentCursor^;
    Inc(CurrentCursor);
    Inc(EltCursor);
  end;
  Inc(FPosition, celt);
  if pceltFetched <> nil then
    pceltFetched^ := celt;
  Result := S_OK;
end;

//  Функция устанавливает позицию в начальное состояние
// =============================================================================
function TFWEnumFormatEtc.Reset: HResult;
begin
  FPosition := 0;
  Result := S_OK;
end;

//  Функция сдвигает текущую позицию на celt элементов
// =============================================================================
function TFWEnumFormatEtc.Skip(celt: Integer): HResult;
begin
  if celt + FPosition > Fcch then
  begin
    FPosition := Fcch;
    Result := S_FALSE;
  end
  else
  begin
    Inc(FPosition, celt);
    Result := S_OK;
  end;
end;

{ TFWDragDropSourceProvider }

//
//  Процедура добавляет новый дескриптор файла/папки к списку уже существующих
//  Финальный список дескрипторов передается эксплореру
//  (или любому другому объекту, реализующему интерфейс IDropTarget)
//  в тот момент, когда он вызовет метод нашего интерфейса EnumFormatEtc
// =============================================================================
procedure TFWDataObject.AddDescriptor(Value: TFileDescriptor);
begin
  Inc(FValues.cItems);
  SetLength(FValues.fgd, FValues.cItems);
  if ShowUIDialog then
    Value.dwFlags := Value.dwFlags or FD_PROGRESSUI
  else
    Value.dwFlags := Value.dwFlags and not FD_PROGRESSUI;
  Value.dwFlags := Value.dwFlags or FD_LINKUI;
  FValues.fgd[FValues.cItems - 1] := Value;
end;

//
//  Процедура добавляет новый файл для копирования
//  к списку существующих дескрипторов
// =============================================================================
procedure TFWDataObject.AddFile(const FileName: string);
begin
  AddFile(FileName, 0);
end;

//
//  Процедура добавляет новый файл для копирования с указанием размера
//  к списку существующих дескрипторов
// =============================================================================
procedure TFWDataObject.AddFile(const FileName: string; FileSize: Int64);
var
  Descriptor: TFileDescriptor;
begin
  ZeroMemory(@Descriptor, SizeOf(TFileDescriptor));
  Descriptor.dwFlags := FD_FILESIZE;
  Descriptor.nFileSizeHigh := FileSize shr 32;
  Descriptor.nFileSizeLow := FileSize and $FFFFFFFF;
  Move(FileName[1], Descriptor.cFileName[0], Length(FileName) * SizeOf(Char));
  AddDescriptor(Descriptor);
end;

//
//  Процедура добавляет новую папку для копирования
//  к списку существующих дескрипторов
// =============================================================================
procedure TFWDataObject.AddFolder(const FileName: string);
var
  Descriptor: TFileDescriptor;
begin
  ZeroMemory(@Descriptor, SizeOf(TFileDescriptor));
  Descriptor.dwFlags := FD_ATTRIBUTES;
  Descriptor.dwFileAttributes := FILE_ATTRIBUTE_DIRECTORY;
  Move(FileName[1], Descriptor.cFileName[0], Length(FileName) * SizeOf(Char));
  AddDescriptor(Descriptor);
end;

//
//  Первоначальная инициализация класса
// =============================================================================
constructor TFWDataObject.Create;
begin
  inherited Create;
  FShowUIDialog := True;
  FSupportedDragEffect := [deCopy];
end;

//
//  IAdviseSink не поддерживается
// =============================================================================
function TFWDataObject.DAdvise(const formatetc: TFormatEtc;
  advf: Integer; const advSink: IAdviseSink;
  out dwConnection: Integer): HResult;
begin
  Result := OLE_E_ADVISENOTSUPPORTED;
end;

//  Метод вызывается при савершении процесса приема данных
// =============================================================================
procedure TFWDataObject.DoEndCopy(CopyStyle: TFWDragEffects);
begin
  if Assigned(FEndCopy) then
    FEndCopy(Self, CopyStyle);
end;

//
//  Метод вызывается при запросе собержимого файлов,
//  ранее добавленных к списку дескрипторов
// =============================================================================
procedure TFWDataObject.DoGetContent(
  FileDescriptor: TFileDescriptor; Data: TStream);
begin
  if Assigned(FContent) then
  begin
    FContent(Self, FileDescriptor, Data);
    Data.Position := 0;
  end;
end;

//
//  IAdviseSink не поддерживается
// =============================================================================
function TFWDataObject.DUnadvise(dwConnection: Integer): HResult;
begin
  Result := OLE_E_ADVISENOTSUPPORTED;
end;

//
//  IAdviseSink не поддерживается
// =============================================================================
function TFWDataObject.EnumDAdvise(
  out enumAdvise: IEnumStatData): HResult;
begin
  Result := OLE_E_ADVISENOTSUPPORTED;
end;

//
//  Приемник запрашивает список поддерживаемых форматов
// =============================================================================
function TFWDataObject.EnumFormatEtc(dwDirection: Integer;
  out enumFormatEtc: IEnumFormatEtc): HResult;
var
  FormatsCount, I, Cursor: Integer;
  NeedHDrop: Boolean;
begin
  enumFormatEtc := nil;
  Result := E_NOTIMPL;
  if dwDirection = DATADIR_GET then
  begin
    NeedHDrop := False;
    for I := 0 to FValues.cItems - 1 do
      if FileExists(PChar(@FValues.fgd[I].cFileName[0])) then
      begin
        NeedHDrop := True;
        Break;
      end;
    FormatsCount := 2 + FValues.cItems;
    if NeedHDrop then
      Inc(FormatsCount);
    SetLength(FFormatEtc, FormatsCount);

    // Говорим что поддерживаем список дескрипторов файлов
    Cursor := 0;
    if NeedHDrop then
    begin
      FFormatEtc[Cursor] := GetFormatEtc(CF_HDROP);
      Inc(Cursor);
    end;
    FFormatEtc[Cursor] := GetFormatEtc(CF_FILEDESCRIPTOR);
    Inc(Cursor);
    FFormatEtc[Cursor] := GetFormatEtc(CF_PREFERREDDROPEFFECT);
    Inc(Cursor);

    // Подготавливаем список содержимого файлов,
    // (соответствующий списку формата CF_FILEDESCRIPTOR), на основании которого
    // сервер будет запрашивать у нас содержимое каждого файла,
    // передавая заполненую нами структуру в формате CF_FILECONTENTS
    for I := Cursor to FormatsCount - 1 do
      FFormatEtc[I] := GetFormatEtc(CF_FILECONTENTS, nil,
        DVASPECT_CONTENT, I - Cursor, TYMED_HGLOBAL or TYMED_ISTREAM);

    // На запрос отдаем интерфейс IEnumFormatEtc,
    // через который внешнее приложение просмотрит весь список форматов
    enumFormatEtc := TFWEnumFormatEtc.Create(@FFormatEtc[0], FormatsCount);
    Result := S_OK;
  end;
end;

//
//  Этот метод не поддерживается,
//  говорим что внешнее приложение запросило у нас правильную структуру
// =============================================================================
function TFWDataObject.GetCanonicalFormatEtc(
  const formatetc: TFormatEtc; out formatetcOut: TFormatEtc): HResult;
begin
  formatetcOut := formatetc;
  Result := DATA_S_SAMEFORMATETC;
end;

//
//  Через данный метод внешнее приложение запрашивает данные у нас
// =============================================================================
function TFWDataObject.GetData(const formatetcIn: TFormatEtc;
  out medium: TStgMedium): HResult;
var
  hMem: HGLOBAL;
  Size: DWORD;
  FileContent: TMemoryStream;
  Stream: IStream;
  FGD: PFileGroupDescriptorEx;
  pDropEffect: PDWORD;
  lpDrop: PDropFiles;
  lpChar: PChar;
  I: Integer;
begin
  // Обязательная инициализация возвращаемой структуры
  medium.tymed := 0;
  medium.hGlobal := 0;
  medium.unkForRelease := nil;
  Result := DV_E_FORMATETC;

  // Проверка - поддерживаем ли мы запрашиваемый формат данных?
  if (IsValidFormat(formatetcIn) = S_OK) then
  begin
    // Да, формат поддерживается -
    // теперь смотрим, какой именно формат был запрошен?


    // Запрос файлов присутствующих на диске
    if formatetcIn.cfFormat = CF_HDROP then
    begin

      // Рассчет размера под данные
      Size := SizeOf(TDropFiles) + 2 * SizeOf(Char);
      for I := 0 to FValues.cItems - 1 do
      begin
        if FileExists(PChar(@FValues.fgd[I].cFileName[0])) then
          Inc(Size, (Length(PChar(@FValues.fgd[I].cFileName[0])) + 1) * SizeOf(Char));
      end;

      if Size > $FFFF then
      begin
        Result := E_FAIL;
        Exit;
      end;

      hMem := GlobalAlloc(GHND, Size);
      if hMem = 0 then
      begin
        Result := E_OUTOFMEMORY;
        Exit;
      end;

      lpDrop := GlobalLock(hMem);
      try
        lpDrop^.pFiles := SizeOf(TDropFiles);
        {$IFDEF UNICODE}
        DWORD(lpDrop^.fWide) := 1;
        {$ENDIF}
        lpChar := PChar(PByte(lpDrop) + SizeOf(TDropFiles));
        for I := 0 to FValues.cItems - 1 do
        begin
          if FileExists(PChar(@FValues.fgd[I].cFileName[0])) then
          begin
            Move(FValues.fgd[I].cFileName[0], lpChar^,
              Length(PChar(@FValues.fgd[I].cFileName[0])) * SizeOf(Char));
            Inc(lpChar, Length(PChar(@FValues.fgd[I].cFileName[0])) + 1);
          end;
        end;
      finally
        GlobalUnlock(hMem);
      end;
      // Выставляем флаг типа передачи данных
      medium.tymed := TYMED_HGLOBAL;
      // и указатель на область памяти с запрошенными данными
      medium.hGlobal := hMem;
      // за освобождение выделенной памяти отвечает приемник
      Result := S_OK;
    end;

    // Произошел запрос списка дескрипторов передаваемых файлов
    if formatetcIn.cfFormat = CF_FILEDESCRIPTOR then
    begin
      // В этом случае данные передаются через medium.hGlobal
      // в виде структуры TFileGroupDescriptor (не TFileGroupDescriptorEx)

      // Рассчитываем требуемый размер памяти
      Size := SizeOf(UINT) + FValues.cItems * SizeOf(TFileDescriptor);
      // Выделяем память
      hMem := GlobalAlloc(GHND, Size);
      if hMem = 0 then
      begin
        Result := E_OUTOFMEMORY;
        Exit;
      end;
      FGD := GlobalLock(hMem);
      try
        // Копируем данные
        CopyMemory(Pointer(FGD), @FValues.cItems, SizeOf(UINT));
        CopyMemory(Pointer(NativeUInt{DWORD}(FGD) + SizeOf(UINT)),
          @FValues.fgd[0], Size - SizeOf(UINT));
      finally
        GlobalUnlock(hMem);
      end;
      // Выставляем флаг типа передачи данных
      medium.tymed := TYMED_HGLOBAL;
      // и указатель на область памяти с запрошенными данными
      medium.hGlobal := hMem;
      // за освобождение выделенной памяти отвечает приемник
      Result := S_OK;
    end;

    // Пришел запрос поддерживаемых стилей при перетаскивании
    if formatetcIn.cfFormat = CF_PREFERREDDROPEFFECT then
    begin
      // В этом случае данные передаются через medium.hGlobal
      // Который указывает на 4-байтовое значение

      hMem := GlobalAlloc(GHND, SizeOf(DWORD));
      if hMem = 0 then
      begin
        Result := E_OUTOFMEMORY;
        Exit;
      end;
      // Запрос поддерживаемых стилей
      pDropEffect := GlobalLock(hMem);
      try
        pDropEffect^ := EffectToDWORD(FSupportedDragEffect);
      finally
        GlobalUnlock(hMem);
      end;
      // Выставляем флаг типа передачи данных
      medium.tymed := TYMED_HGLOBAL;
      // и указатель на область памяти с запрошенными данными
      medium.hGlobal := hMem;
      // за освобождение выделенной памяти отвечает приемник
      Result := S_OK;
    end;

    // Пришел запрос о содержимом файла
    if formatetcIn.cfFormat = CF_FILECONTENTS then
    begin
      // Смотрим какой номер файла запрашивается (валидный или нет)
      if formatetcIn.lindex < 0 then
      begin
        Result := DV_E_LINDEX;
        Exit;
      end;
      // Если мы регистрировали данный файл - запрашиваем его содержимое
      FileContent := TMemoryStream.Create;
      DoGetContent(FValues.fgd[formatetcIn.lindex], FileContent);
      medium.tymed := TYMED_ISTREAM;
      // Отдаем содеримое через TStreamAdapter и назначаем его владельцем
      // созданного TMemoryStream. Как только количество ссылок на
      // TStreamAdapter станет равно нулю, он разрушится и разрушит FileContent
      stream := TStreamAdapter.Create(FileContent, soOwned);
      medium.stm := Pointer(stream);
      // вот тут явный гавнокод из-за инкремена рефов
      // должен быть более правильный способ, но банально лениво его искать :)
      stream._AddRef;
      Result := S_OK;
    end;
  end;
end;

//
//  Не поддерживается
// =============================================================================
function TFWDataObject.GetDataHere(const formatetc: TFormatEtc;
  out medium: TStgMedium): HResult;
begin
  Result := E_NOTIMPL;
end;

//
//  Функция предназначена для проверки,
//  поддерживает ли наш интерфейс запрашиваемый формат данных
// =============================================================================
function TFWDataObject.IsValidFormat(
  formatetc: TFormatEtc): HResult;
var
  I, A: Integer;
  TymedFound: Boolean;

  function CheckTymed(Value: Integer): Boolean;
  begin
    Result := True;
    if (FFormatEtc[I].tymed and Value) = Value then
      Result := (formatetc.tymed and Value) = Value;
  end;

begin
  Result := DV_E_FORMATETC;
  for I := 0 to Length(FFormatEtc) - 1 do
    if FFormatEtc[I].dwAspect = FormatEtc.dwAspect then
      if FFormatEtc[I].cfFormat = FormatEtc.cfFormat then
        if FFormatEtc[I].lindex = FormatEtc.lindex then
        begin
          Result := DV_E_TYMED;
          TymedFound := False;
          for A := 0 to 6 do
            if CheckTymed(TYMED_ARRAY[A]) then
            begin
              TymedFound := True;
              Break;
            end;
          if TymedFound then
            Result := S_OK;
          Break;
        end;
end;

//
//  Через данный метод внешнее приложение может проверить,
//  может ли оно получить определенный тип данных
// =============================================================================
function TFWDataObject.QueryGetData(
  const formatetc: TFormatEtc): HResult;
begin
  Result := IsValidFormat(formatetc);
end;

//
//  Через данный метод внешнее приложение сигнализирует нам
//  о завершении приема данных
// =============================================================================
function TFWDataObject.SetData(const formatetc: TFormatEtc;
  var medium: TStgMedium; fRelease: BOOL): HResult;
var
  pEffect: PInteger;
begin
  Result := DV_E_FORMATETC;
  try
    if ((formatetc.cfFormat = CF_PERFORMEDDROPEFFECT) or
      (formatetc.cfFormat = CF_PASTESUCCEEDED)) and
      (formatetc.tymed = TYMED_HGLOBAL) then
    begin
      pEffect := GlobalLock(medium.hGlobal);
      if pEffect = nil then
      begin
        Result := E_OUTOFMEMORY;
        Exit;
      end;
      try
        DoEndCopy(DWORDToEffect(pEffect^, [deNone..deScroll]));
      finally
        GlobalUnlock(medium.hGlobal);
      end;
      Result := S_OK;
    end;
  finally
    if fRelease then
      ReleaseStgMedium(medium);
  end;
end;

//
//  Выставляем флаг, отображать диалог копирования или нет
// =============================================================================
procedure TFWDataObject.SetShowUIDialog(const Value: Boolean);
var
  I: Integer;
begin
  FShowUIDialog := Value;
  for I := 0 to Integer(FValues.cItems) - 1 do
    if Value then
      FValues.fgd[I].dwFlags := FValues.fgd[I].dwFlags and not FD_PROGRESSUI
    else
      FValues.fgd[I].dwFlags := FValues.fgd[I].dwFlags or FD_PROGRESSUI;
end;

{ TFWDropTarget }

//
//  Очищаем список дескрипторов
// =============================================================================
procedure TFWDropTarget.ClearFiles;
begin
  FFiles.cItems := 0;
  SetLength(FFiles.fgd, 0);
end;

//
//  Базовый конструктор класса
// =============================================================================
constructor TFWDropTarget.Create;
begin
  inherited Create;
  FDE := [deCopy];
end;

//
//  Вызов внешнего обработчика OnDragOver - спрашиваем реакцию извне
// =============================================================================
procedure TFWDropTarget.DoDragOver(pt: TPoint;
  Files: TFileGroupDescriptorEx; DragState: TDragState;
  var DragEffect: TFWDragEffects);
begin
  if Assigned(FDragOver) then
    FDragOver(Self, pt, Files, DragState, DragEffect);
end;

//
//  Вызов внешнего обработчика OnSetContent - отдаем результат программисту
// =============================================================================
procedure TFWDropTarget.DoSetContent(const FileDescriptor: TFileDescriptor;
  Data: TStream; DragEffect: TFWDragEffects);
begin
  if Assigned(FContent) then
  begin
    Data.Position := 0;
    FContent(Self, FileDescriptor, Data, DragEffect);
  end;
end;

//
//  Получили уведомление о начале операции перетаскивания над нашим контролом
// =============================================================================
function TFWDropTarget.DragEnter(const dataObj: IDataObject;
  grfKeyState: Integer; pt: TPoint; var dwEffect: Integer): HResult;
var
  DragEffect: TFWDragEffects;
begin
  DragEffect := DWORDToEffect(dwEffect, SupportedDragEffect);
  // получаем список перетаскиваемых объектов
  ExtractFileNames(dataObj, pt);
  // отправляем их на внешку программисту
  DoDragOver(pt, FFiles, dsDragEnter, DragEffect);
  // говорим о принятом решении
  FDE := DragEffect;
  dwEffect := EffectToDWORD(DragEffect);
  Result := S_OK;
end;

//
//  Получили уведомление о завершении операции перетаскивания
// =============================================================================
function TFWDropTarget.DragLeave: HResult;
var
  DragEffect: TFWDragEffects;
begin
  ClearFiles;
  DragEffect := [deNone];
  // отправляем уведомление на внешку
  DoDragOver(Point(-1, -1), FFiles, dsDragLeave, DragEffect);
  Result := S_OK;
end;

//
//  Получили уведомление о происходящей операции перетаскивания
// =============================================================================
function TFWDropTarget.DragOver(grfKeyState: Integer; pt: TPoint;
  var dwEffect: Integer): HResult;
var
  DragEffect: TFWDragEffects;
begin
  DragEffect := DWORDToEffect(dwEffect, SupportedDragEffect);
  // спрашиваем программиста - как на нее будем реагировать
  DoDragOver(pt, FFiles, dsDragMove, DragEffect);
  FDE := DragEffect;
  dwEffect := EffectToDWORD(DragEffect);
  Result := S_OK;
end;

//
//  Получили уведомление о готовности отправки данных о перетащенных данных
// =============================================================================
function TFWDropTarget.Drop(const dataObj: IDataObject; grfKeyState: Integer;
  pt: TPoint; var dwEffect: Integer): HResult;

  function IsFile(Value: TFileDescriptor): Boolean;
  begin
    Result := (Value.dwFileAttributes and FILE_ATTRIBUTE_DIRECTORY) = 0;
  end;

var
  F: TMemoryStream;
  I: Integer;
  AFormatEtc: TFormatEtc;
  medium: TStgMedium;
  stat: TStatStg;
  stream: IStream;
  cbRead, cbWritten: LargeUInt;
  pEffect: PInteger;
  DataPresent: Boolean;
  P: Pointer;
begin
  // витуальные данные будем сохранять во временном стриме
  F := TMemoryStream.Create;
  try
    // смотрим, работаем ли мы с виртуальными данными (файла на диске нема)
    if FVirtualFiles then
    begin
      for I := 0 to FFiles.cItems - 1 do
      begin
        F.Clear;
        // так себе проверка, но все-же - смотрим чтобы это небыло директорией
        // момент спорный и не гарантирующий практически вообще ничего
        if IsFile(FFiles.fgd[I]) then
        begin
          DataPresent := False;
          // проверяем, можем ли мы получить данные посредством IStream
          AFormatEtc := GetFormatEtc(
            CF_FILECONTENTS, nil, DVASPECT_CONTENT, I,
            TYMED_ISTREAM);
          if dataObj.QueryGetData(AFormatEtc) = S_OK then
          begin
            if dataObj.GetData(AFormatEtc, medium) = S_OK then
            try
              // да можем - вытягиваем их
              DataPresent := True;
              stream := TStreamAdapter.Create(F, soReference);
              IStream(medium.stm).Stat(stat, STATFLAG_NONAME);
              IStream(medium.stm).CopyTo(
                stream, stat.cbSize, cbRead, cbWritten);
            finally
              ReleaseStgMedium(medium);
            end;
          end;
          // проверка, пришли ли данные через IStream?
          if not DataPresent then
          begin
            // если не - пытаемся их забрать через глобалку
            AFormatEtc := GetFormatEtc(
              CF_FILECONTENTS, nil, DVASPECT_CONTENT, I,
              TYMED_HGLOBAL);
            if dataObj.QueryGetData(AFormatEtc) = S_OK then
            begin
              if dataObj.GetData(AFormatEtc, medium) = S_OK then
              try
                P := GlobalLock(medium.hGlobal);
                try
                  F.Size := GlobalSize(medium.hGlobal);
                  CopyMemory(F.Memory, P, F.Size);
                finally
                  GlobalUnlock(medium.hGlobal);
                end;
              finally
                ReleaseStgMedium(medium);
              end;
            end;
          end;
        end;
        // уведомляем программиста о приеме данных
        DoSetContent(FFiles.fgd[I], F,
          DWORDToEffect(dwEffect, SupportedDragEffect));
      end;
      Result := S_OK;
    end
    else
    begin
      // ну а если переданные файлы не виртуальные
      // то просто передаем на внешку их путь
      for I := 0 to FFiles.cItems - 1 do
      begin
        F.Clear;
        DoSetContent(FFiles.fgd[I], F,
          DWORDToEffect(dwEffect, SupportedDragEffect));
      end;
      Result := S_OK;
    end;
  finally
    F.Free;
  end;

  // через глобалку возвращаем результат операции посредством CF_PASTESUCCEEDED
  medium.tymed := TYMED_HGLOBAL;
  medium.unkForRelease := nil;
  medium.hGlobal := GlobalAlloc(GHND, 4);
  if medium.hGlobal = 0 then
  begin
    Result := E_OUTOFMEMORY;
    Exit;
  end;
  try
    pEffect := GlobalLock(medium.hGlobal);
    if pEffect <> nil then
    try
      pEffect^ := dwEffect;
    finally
      GlobalUnlock(medium.hGlobal);
    end;
    AFormatEtc := GetFormatEtc(CF_PASTESUCCEEDED);
    dataObj.SetData(AFormatEtc, medium, False);
  finally
    ReleaseStgMedium(medium);
  end;
end;

//
//  Вспомогалочка, через интерфейс получаем указатель на экземпляр класса
// =============================================================================
function TFWDropTarget.GetInstance: TFWDropTarget;
begin
  Result := Self;
end;

//
//  Получаем параметры переданных файлов
// =============================================================================
procedure TFWDropTarget.ExtractFileNames(const dataObj: IDataObject;
  pt: TPoint);
var
  I: Integer;
  AFormatEtc: TFormatEtc;
  medium: TStgMedium;
  FGD: PFileGroupDescriptorEx;
  SR: TSearchRec;
begin
  ClearFiles;

  // для начала смотрим инфу о реальных файлах, переданных посредством CF_HDROP
  AFormatEtc := GetFormatEtc(CF_HDROP);
  if dataObj.QueryGetData(AFormatEtc) = S_OK then
  begin
    if dataObj.GetData(AFormatEtc, medium) = S_OK then
    try
      FVirtualFiles := False;
      FFiles.cItems := DragQueryFile(medium.hGlobal, DWORD(-1), nil, 0);
      SetLength(FFiles.fgd, FFiles.cItems);
      ZeroMemory(@FFiles.fgd[0], FFiles.cItems * SizeOf(TFileDescriptor));
      for I := 0 to FFiles.cItems - 1 do
      begin
        DragQueryFile(medium.hGlobal, I, @FFiles.fgd[I].cFileName[0], MAX_PATH);
        // CF_HDROP применяется только для реально существующих на диске файлов
        if FindFirst(FFiles.fgd[I].cFileName, faAnyFile, SR) = 0 then
        try
          FFiles.fgd[I].dwFlags := FD_ACCESSTIME or FD_ATTRIBUTES or
            FD_CREATETIME or FD_FILESIZE or FD_WRITESTIME;
          FFiles.fgd[I].dwFileAttributes := SR.FindData.dwFileAttributes;
          FFiles.fgd[I].ftCreationTime := SR.FindData.ftCreationTime;
          FFiles.fgd[I].ftLastAccessTime := SR.FindData.ftLastAccessTime;
          FFiles.fgd[I].ftLastWriteTime := SR.FindData.ftLastWriteTime;
          FFiles.fgd[I].nFileSizeHigh := SR.FindData.nFileSizeHigh;
          FFiles.fgd[I].nFileSizeLow := SR.FindData.nFileSizeLow;
        finally
          FindClose(SR);
        end;
      end;
      DragFinish(Medium.hGlobal);
      Exit;
    finally
      ReleaseStgMedium(medium);
    end;
  end;

  // после чего смотрим есть ли вируальные файлы
  AFormatEtc := GetFormatEtc(CF_FILEDESCRIPTOR);
  if dataObj.QueryGetData(AFormatEtc) = S_OK then
  begin
    if dataObj.GetData(AFormatEtc, medium) = S_OK then
    try
      FVirtualFiles := True;
      FGD := GlobalLock(medium.hGlobal);
      try
        FFiles.cItems := FGD^.cItems;
        SetLength(FFiles.fgd, FFiles.cItems);
        CopyMemory(@FFiles.fgd[0], Pointer(DWORD(FGD) + 4),
          FFiles.cItems * SizeOf(TFileDescriptor));
      finally
        GlobalUnlock(medium.hGlobal);
      end;
    finally
      ReleaseStgMedium(medium);
    end;
  end;
end;

//
//  Регистрируем окно, получающее уведомления
// =============================================================================
procedure TFWDropTarget.RegisterDragDrop(DropTarget: TWinControl);
begin
  RevokeDragDrop;
  FDropTarget := DropTarget;
  ActiveX.RegisterDragDrop(FDropTarget.Handle, Self);
end;

//
//  Снимаем окно с регистрации
// =============================================================================
procedure TFWDropTarget.RevokeDragDrop;
begin
  if FDropTarget <> nil then
    ActiveX.RevokeDragDrop(FDropTarget.Handle);
  FDropTarget := nil;
end;

{ TFWDragDropSource }

//
//  Уведомляем программиста о завершении операции перетаскивания
// =============================================================================
procedure TFWDragDropSourceProvider.DoDropEnd;
begin
  if Assigned(FDropEnd) then
    FDropEnd(Self);
end;

//
//  Инициализируем операцию перетаскивания
// =============================================================================
function TFWDragDropSourceProvider.ExecuteDragDrop(
  const DragEffect: TFWDragEffects): HRESULT;
var
  dwEffect: Integer;
begin
  FSupportedDragEffect := DragEffect;
  Result := DoDragDrop(Self, Self, EffectToDWORD(DragEffect), dwEffect);
end;

//
//  Уведомляем о виде курсора и говорим что всегда юзаем курсор по умолчанию
//  (это была первая ошибка в предыдущем варианте)
// =============================================================================
function TFWDragDropSourceProvider.GiveFeedback(dwEffect: Integer): HResult;
begin
  Result := DRAGDROP_S_USEDEFAULTCURSORS;
end;

//
//  Уведомляем о статусе операции
//  (это была вторая ошибка в предыдущем варианте, я возвращал не S_OK а S_FALSE)
// =============================================================================
function TFWDragDropSourceProvider.QueryContinueDrag(fEscapePressed: BOOL;
  grfKeyState: Integer): HResult;
begin
  Result := S_OK;
  if fEscapePressed then
  begin
    DoDropEnd;
    Result := DRAGDROP_S_CANCEL;
    Exit;
  end;
  if ((grfKeyState and MK_LBUTTON) = 0) and
    ((grfKeyState and MK_RBUTTON) = 0) then
  begin
    DoDropEnd;
    Result := DRAGDROP_S_DROP;
  end;
end;

{ TFWCopyDataProvider }

//
//  Чистим буфер обмена
// =============================================================================
procedure TFWCopyDataProvider.ClearClipboard;
begin
  OleSetClipboard(nil);
end;

//
//  Весь обьект копируем в буфер обмена
// =============================================================================
function TFWCopyDataProvider.CopyToClibboard: HRESULT;
begin
  Result := OleSetClipboard(Self);
end;

//
//  Ищем PIDL обьекта для вставки данных и вызываем на нем Shell 'Paste'
// =============================================================================
function TFWCopyDataProvider.PasteToFolder(const Path: string): HRESULT;
var
  Desktop, ShellFolder: IShellFolder;
  PathPIDL: PItemIDList;
  FilePIDL: array [0..1] of PItemIDList;
  pchEaten, Attr: Cardinal;
  ICMenu: IContextMenu;
  cmICI: TCMInvokeCommandInfo;
begin
  Result := 0;
  try
    OleCheck(SHGetDesktopFolder(Desktop));
    OleCheck(SHGetSpecialFolderLocation(0, CSIDL_DRIVES, PathPIDL));
    OleCheck(Desktop.BindToObject(PathPIDL, nil,
      IID_IShellFolder, ShellFolder));
    OleCheck(ShellFolder.ParseDisplayName(0, nil, StringToOleStr(Path),
      pchEaten, FilePIDL[0], Attr));
    OleCheck(ShellFolder.GetUIObjectOf(0, 1, FilePIDL[0],
      IID_IContextMenu, nil, ICMenu));
    ZeroMemory(@cmICI, SizeOf(TCMInvokeCommandInfo));
    cmICI.cbSize := SizeOf(TCMInvokeCommandInfo);
    cmICI.lpVerb := 'Paste';
    cmICI.nShow := SW_SHOWNORMAL;
    OleCheck(ICMenu.InvokeCommand(cmICI));
    ClearClipboard;
  except
    on E: EOleSysError do
    begin
      Result := E.ErrorCode;
      Exit;
    end;
    on E: Exception do
      raise;
  end;
end;

{ TFWPasteDataProvider }

//
//  Получаем список файлов для вставки из буфера обмена
// =============================================================================
class function TFWPasteDataProvider.GetFileListFromClipboard: TStringList;
var
  FmtEtc: TFormatEtc;
  Medium: TStgMedium;
  dataObj: IDataObject;
  I: Integer;
  FileName: string;
begin
  Result := TStringList.Create;
  if not IsClipboardFormatAvailable(CF_HDROP) then Exit;
  if Failed(OleGetClipboard(dataObj)) then Exit;
  FmtEtc := GetFormatEtc(CF_HDROP);
  if Failed(DataObj.GetData(FmtEtc, Medium)) then Exit;
  try
    for I := 0 to DragQueryFile(Medium.hGlobal, DWORD(-1), nil, 0) - 1 do
    begin
      SetLength(FileName, DragQueryFile(Medium.hGlobal, I, nil, 0));
      DragQueryFile(Medium.hGlobal, I, PChar(FileName), Length(FileName) + 1);
      Result.Add(FileName);
    end;
    DragFinish(Medium.hGlobal);
  finally
    ReleaseStgMedium(Medium);
  end;
end;

//
//  Проверяем - есть ли доступные для вставки файлы в буфере обмена
// =============================================================================
class function TFWPasteDataProvider.IsFileListAvailable: Boolean;
var
  FmtEtc: TFormatEtc;
  Medium: TStgMedium;
  dataObj: IDataObject;
begin
  Result := False;
  if Failed(OleGetClipboard(dataObj)) then Exit;
  FmtEtc := GetFormatEtc(CF_HDROP);
  if Failed(DataObj.GetData(FmtEtc, Medium)) then Exit;
  try
    Result := DragQueryFile(Medium.hGlobal, DWORD(-1), nil, 0) > 0;
    DragFinish(Medium.hGlobal);
  finally
    ReleaseStgMedium(Medium);
  end;
end;

initialization
  // регистрация всей этой белиберды
  OleInitialize(nil);
  CF_FILEDESCRIPTOR := RegisterClipboardFormat(CFSTR_FILEDESCRIPTOR);
  CF_FILECONTENTS := RegisterClipboardFormat(CFSTR_FILECONTENTS);
  CF_PREFERREDDROPEFFECT := RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT);
  CF_PERFORMEDDROPEFFECT := RegisterClipboardFormat(CFSTR_PERFORMEDDROPEFFECT);
  CF_PASTESUCCEEDED := RegisterClipboardFormat(CFSTR_PASTESUCCEEDED);

finalization
  // финализация
  // мы же не хотим держать в буфере обмена что-то,
  // что будет не доступно снаружи после завершения нашего приложения :)
  OleFlushClipboard;
  OleUninitialize;

end.

