unit fpsRichTextCtrls;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  RichMemo,
  fpSpreadsheet, fpsTypes, fpsUtils, fpSpreadsheetCtrls;

type
  TsCustomCellRichMemo = class(TRichMemo)
  private
    procedure ApplyFont(ATextStart, ATextLength: Integer; AFont: TsFont; ApplySize: Boolean);
    procedure ApplyRichParams(ACell: PCell);
    function ExtractRichTextParams(ACell: PCell): TsRichTextParams;
  protected
  public
    procedure CellToMemo(ACell: PCell);
    procedure MemoToCell(AWorksheet: TsWorksheet; ARow, ACol: Cardinal);
  end;
              (*
  TsRichCellEdit = class(TRichMemo, IsSpreadsheetControl)
  private
    FOldText: String;
    FRefocusing: TObject;
    FRefocusingCol, FRefocusingRow: Cardinal;
    FRefocusingSelStart: Integer;
    FRichTextParams: TsRichTextParams;
    FWorkbookSource: TsWorkbookSource;
    function GetSelectedCell: PCell;
    function GetWorkbook: TsWorkbook;
    function GetWorksheet: TsWorksheet;
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
    
  protected
    procedure ApplyFont(ATextStart, ATextLen: Integer; AFont: TsFont; ApplySize: boolean);
    procedure ApplyRichParams;
    function ExtractRichTextParams: TsRichTextParams;
    function CanEditCell(ACell: PCell): Boolean; overload;
    function CanEditCell(ARow, ACol: Cardinal): Boolean; overload;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure ShowCell(ACell: PCell); virtual;
    
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ListenerNotification(AChangedItems: TsNotificationItems;
      AData: Pointer = nil);
    procedure RemoveWorkbookSource;
    property SelectedCell: PCell read GetSelectedCell;
    property Workbook: TsWorkbook read GetWorkbook;
    property Worksheet: TsWorksheet read GetWorksheet;
    
  published
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
    
  end;
  *)
  
implementation

uses
  fpsVisualUtils;

function Convert_sFont_to_FontParams(AFont: TsFont): TFontParams;
begin
  Result := Default(TFontParams);
  Result.Name := AFont.FontName;
  Result.Style := Convert_sFontStyle_to_FontStyle(AFont.Style);
  Result.Color := Convert_sColor_to_Color(AFont.Color);
  case AFont.Position of
    fpSuperscript: Result.VScriptPos :=  vpSuperscript;
    fpSubscript: Result.VScriptPos := vpSubscript;
    fpNormal:  Result.VScriptPos := vpNormal;
  end;
  Result.Size := round(AFont.Size)
end;  

function Convert_VScriptPos_to_sFontPosition(APos: TVScriptPos): TsFontPosition;
begin
  case APos of
    vpNormal: Result := fpNormal;
    vpSubscript: Result := fpSubscript;
    vpSuperscript: Result := fpSuperscript;
  end;
end;  

function SameFontParams(fp1, fp2: TFontParams): Boolean;
const
  EPS = 1E-6;
begin
  Result := 
    SameText(fp1.Name, fp2.Name) and 
    (abs(fp1.Size - fp2.Size) < EPS) and
    (fp1.Color = fp2.Color) and
    (fp1.Style = fp2.Style) and
    (fp1.HasBkClr = fp2.HasBkClr) and
    (fp1.BkColor = fp2.BkColor) and
    (fp1.VScriptPos = fp2.VScriptPos);
end;
  
// -----------------------------------------------------------------------------
//                            TsCustomCellRichMemo
// -----------------------------------------------------------------------------

procedure TsCustomCellRichMemo.ApplyFont(ATextStart, ATextLength: Integer; 
  AFont: TsFont; ApplySize: boolean);
var
  fp: TFontParams;
begin
  fp := Convert_sFont_to_FontParams(AFont);
  if not ApplySize then
    fp.Size := 10; //  FIX ME  !!!!            
  SetTextAttributes(ATextStart-1, ATextLength, fp);  
end;
  
procedure TsCustomCellRichMemo.ApplyRichParams(ACell: PCell);
var
  i, last: Integer;
  rtp: TsRichTextParam;
  startPos, count: Integer;
  fnt: TsFont;
  cellRichTextParams: TsRichTextParams;
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  cellText: String;
begin
  if (ACell = nil) or (ACell^.ContentType <> cctUTF8String) then
    exit;
  
  cellText := ACell^.UTF8StringValue;
  if Length(cellText) = 0 then 
    exit;

  worksheet := TsWorksheet(ACell^.Worksheet);
  workbook := TsWorkbook(worksheet.Workbook);
  cellRichTextParams := ACell^.RichTextParams;
  
  if Length(cellRichTextParams) = 0 then
  begin
    fnt := worksheet.ReadCellFont(ACell);
    count := Length(cellText);
    ApplyFont(1, count, fnt, true);
    exit;
  end;
  
  i := 0;
  last := High(cellRichTextParams);
  startPos := 1;
  
  rtp := cellRichTextParams[0];
  if rtp.FirstIndex > 1 then
  begin
    fnt := worksheet.ReadCellFont(ACell);
    count := rtp.FirstIndex;
    ApplyFont(startPos, count, fnt, true);
  end;
  
  while i <= last do
  begin
    rtp := cellRichTextParams[i];
    startPos := rtp.FirstIndex;
    if i < last then
      count := cellRichTextParams[i+1].FirstIndex
    else
      count := Length(cellText) - startPos + 1;
    fnt := workbook.GetFont(rtp.FontIndex);
    ApplyFont(startPos, count, fnt, true);
    inc(i);
  end;
  
{  
  rtp := cellRichTextParams[0];
  count := rtp.FirstIndex;
  while i < last do
  begin
    ApplyFont(startpos, count, fnt, false);
    rtp := cellRichTextparams[i];
    startPos := rtp.FirstIndex;
    fnt := workbook.GetFont(rtp.FontIndex);
    count := cellRichTextParams[i+1].FirstIndex;
    inc(i);
  end;
  count := Length(Lines.Text) - startPos;
  ApplyFont(startPos, count, fnt, false);
  }
end;

{@@ Moves text and formatting attributes from the cell record (to which ACell
    points) to the RichMemo. }
procedure TsCustomCellRichMemo.CellToMemo(ACell: PCell);
var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  hideFormula: Boolean;
  s: String;
begin
  if ACell = nil then
  begin
    Clear;
    exit;
  end;
  
  worksheet := TsWorksheet(ACell^.Worksheet);
  workbook := TsWorkbook(worksheet.Workbook);
  
  hideformula := worksheet.IsProtected and (spCells in worksheet.Protection) and
    (cpHideFormulas in worksheet.ReadCellProtection(ACell));
  s := worksheet.ReadFormulaAsString(ACell, true);
  if (s <> '') then begin
    if hideformula then
      s := '(Formula hidden)'
    else
      if s[1] <> '=' then s := '=' + s;
    Lines.Text := s;
  end else
    case ACell^.ContentType of
      cctNumber:
        Lines.Text := FloatToStr(ACell^.NumberValue);
      cctDateTime:
        if ACell^.DateTimeValue < 1.0 then        // Time only
          Lines.Text := FormatDateTime('tt', ACell^.DateTimeValue)
        else
        if frac(ACell^.DateTimeValue) = 0 then    // Date only
          Lines.Text := FormatDateTime('ddddd', ACell^.DateTimevalue)
        else                                      // both
          Lines.Text := FormatDateTime('c', ACell^.DateTimeValue);
      cctUTF8String:
        begin
          Lines.Text := ACell^.UTF8StringValue;
          ApplyRichParams(ACell);
        end;
      else
        Lines.Text := worksheet.ReadAsText(ACell);
    end;
end;

{@@ Converts the FontParams of the RichTextMemo control to the RichTextParams
  array used by the cell record to which ACell points. 
  It can be safely assumed that ACell is not nil.}
function TsCustomCellRichMemo.ExtractRichTextParams(ACell: PCell): TsRichTextParams;
var
  i: Integer;
  fnt: TsFont;
  fp, prevfp: TFontParams;
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  s: String;
  rtp: TsRichTextParam;
begin
  SetLength(Result, 0);
  
  if ACell^.ContentType <> cctUTF8String then
    exit;
  
  s := Lines.Text;
  if s = '' then
    exit;
  
  worksheet := TsWorksheet(ACell^.Worksheet);
  workbook := TsWorkbook(worksheet);
  
  fnt := worksheet.ReadCellFont(ACell);
  fp := Convert_sFont_to_FontParams(fnt);
  prevfp := fp;
  
  for i := 1 to Length(s) do
  begin
    GetTextAttributes(i, fp);
    if SameFontParams(fp, prevfp) or (i = Length(s)) then
    begin
      rtp.FirstIndex := i;
      rtp.FontIndex := workbook.AddFont(
        fp.Name, 
        fp.Size, 
        Convert_FontStyle_to_sFontStyle(fp.Style), 
        Convert_Color_to_sColor(fp.Color),
        Convert_VScriptPos_to_sFontPosition(fp.VScriptPos)
      );
      SetLength(Result, Length(Result)+1);
      Result[High(Result)] := rtp;
      prevfp := fp;
    end;
  end;
end;

{@@ Writes the memo back to the cell }
procedure TsCustomCellRichMemo.MemoToCell(AWorksheet: TsWorkSheet; ARow, ACol: Cardinal);
var
  s: String;
  cell: PCell;
begin
  cell := AWorksheet.GetCell(ARow, ACol);
  s := Lines.Text;
  if s = '' then 
    AWorksheet.WriteBlank(cell)
  else
  if (s <> '') and (s[1] = '=') then
    AWorksheet.WriteFormula(cell, Copy(s, 2, Length(s)), true)
  else
  if (cell^.ContentType = cctUTF8String) then
    AWorksheet.WriteText(cell, s, ExtractRichTextParams(cell))
  else
    AWorksheet.WriteCellValueAsString(cell, s);
end;

(*
// Writes the changed cell text back to the cell
procedure TsCustomCellRichMemo.Change;
var
  s: String;
  rtp: TsRichTextParams;
begin
  s := Lines.Text;
  if (s <> '') and (s[1] = '=') then
    Worksheet.WriteFormula(FCell, Copy(s, 2, Length(s)), true)
  else
  begin
    if cell^.ContentType = cctUTF8String then
      Worksheet.WriteText(FCell, s, ExtractRichTextParams)
    else
      Worksheet.WriteCellValueAsString(FCell, s);
  end;
end;

function TsCustomCellRichMemo.ExtractRichTextParams: TsRichTextParams;
var
  i: Integer;
  fnt: TsFont;
  fp, prevfp: TFontParams;
  s: String;
  rtp: TsRichTextParam;
begin
  SetLength(Result, 0);
  
  if FCell^.ContentType <> cctUTF8String then
    exit;
  
  s := Lines.Text;
  
  fnt := Worksheet.ReadCellFont(FCell);
  fp := Convert_sFont_to_FontParams(fnt);
  prevfp := fp;
  
  for i := 1 to Length(s) do
  begin
    GetTextAttributes(i, fp);
    if SameFontParams(fp, prevfp) or (i = Length(s)) then
    begin
      rtp.FirstIndex := i;
      rtp.FontIndex := Workbook.AddFont(
        fp.Name, 
        fp.Size, 
        Convert_FontStyle_to_sFontStyle(fp.Style), 
        Convert_Color_to_sColor(fp.Color),
        Convert_VScriptPos_to_sFontPosition(fp.VScriptPos)
      );
      SetLength(Result, Length(Result)+1);
      Result[High(Result)] := rtp;
      prevfp := fp;
    end;
  end;
end;

function TsCustomCellRichMemo.GetWorkbook: TsWorkbook;
var
  sheet: TsWorksheet;
begin
  sheet := GetWorksheet;
  if sheet <> nil then
    Result := sheet.Workbook
  else
    Result := nil;
end;

function TsCustomCellRichMemo.GetWorksheet: TsWorksheet;
begin
  if FCell <> nil then
    Result := TsWorksheet(FCell^.Worksheet)
  else
    Result := nil;
end;

procedure TsCustomCellRichMemo.SetCell(ACell: PCell);
var
  hideformula: Boolean;
  s: String;
begin
  if ACell <> nil then
  begin
    FCell := ACell;
    hideformula := Worksheet.IsProtected and (spCells in Worksheet.Protection) and
      (cpHideFormulas in Worksheet.ReadCellProtection(ACell));
    s := Worksheet.ReadFormulaAsString(ACell, true);
    if (s <> '') then begin
      if hideformula then
        s := '(Formula hidden)'
      else
        if s[1] <> '=' then s := '=' + s;
      Lines.Text := s;
    end else
      case ACell^.ContentType of
        cctNumber:
          Lines.Text := FloatToStr(ACell^.NumberValue);
        cctDateTime:
          if ACell^.DateTimeValue < 1.0 then        // Time only
            Lines.Text := FormatDateTime('tt', ACell^.DateTimeValue)
          else
          if frac(ACell^.DateTimeValue) = 0 then    // Date only
            Lines.Text := FormatDateTime('ddddd', ACell^.DateTimevalue)
          else                                      // both
            Lines.Text := FormatDateTime('c', ACell^.DateTimeValue);
        cctUTF8String:
          begin
            Lines.Text := ACell^.UTF8StringValue;
            ApplyRichParams(ACell^.RichTextParams);
          end;
        else
          Lines.Text := Worksheet.ReadAsText(ACell);
      end;
  end else
  begin
    Clear;
    FCell := nil;
  end;
end;
       *)
    (*
// -----------------------------------------------------------------------------
//                         TsRichCellEdit
// -----------------------------------------------------------------------------

constructor TsRichCellEdit.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsRichCellEdit. 
  Removes itself from the WorkbookSource's listener list.
-------------------------------------------------------------------------------}
destructor TsRichCellEdit.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

procedure TsRichCellEdit.ApplyFont(ATextStart, ATextLen: Integer; AFont: TsFont; 
  ApplySize: boolean);
var
  tp: TFontParams;
begin
  tp.Name := AFont.FontName;
  tp.Style := Convert_sFontStyle_to_FontStyle(AFont.Style);
  tp.Color := Convert_sColor_to_Color(AFont.Color);
  case AFont.Position of
    fpSuperscript: tp.VScriptPos :=  vpSuperscript;
    fpSubscript: tp.VScriptPos := vpSubscript;
    fpNormal:  tp.VScriptPos := vpNormal;
  end;
  if ApplySize then
    tp.Size := round(AFont.Size)
  else
    tp.Size := 10;  // fix me
  SetTextAttributes(ATextStart, ATextLen, tp);  
end;
  
procedure TsRichCellEdit.ApplyRichParams;
var
  i, last: Integer;
  rtp: TsRichTextParam;
  startPos, count: Integer;
  fnt: TsFont;
begin
  fnt := Worksheet.ReadCellFont(SelectedCell);
  startPos := 0;
  if Length(FRichTextParams) = 0 then
  begin
    count := Length(FRichTextParams);
    ApplyFont(0, count, fnt, false);
    exit;
  end;

  i := 0;
  rtp := FRichTextParams[0];
  count := rtp.FirstIndex;
  last := High(FRichTextParams);
  while i < last do
  begin
    ApplyFont(startpos, count, fnt, false);
    rtp := FRichTextparams[i];
    startPos := rtp.FirstIndex;
    fnt := Workbook.GetFont(rtp.FontIndex);
    count := FRichTextParams[i+1].FirstIndex;
    inc(i);
  end;
  count := Length(Lines.Text) - startPos;
  ApplyFont(startPos, count, fnt, false);
end;

function TsRichCellEdit.CanEditCell(ACell: PCell): Boolean;
begin
  if Worksheet.IsMerged(ACell) then
    ACell := Worksheet.FindMergeBase(ACell);
  Result := not (
    Worksheet.IsProtected and
    (spCells in Worksheet.Protection) and
    ((ACell = nil) or (cpLockcell in Worksheet.ReadCellProtection(ACell)))
  );
end;

function TsRichCellEdit.CanEditCell(ARow, ACol: Cardinal): Boolean;
var
  cell: PCell;
begin
  cell := Worksheet.FindCell(ARow, ACol);
  Result := CanEditCell(cell);
end;

function TsRichCellEdit.ExtractRichTextParams: TsRichTextParams;
begin
  Result := [];
end;

function TsRichCellEdit.GetSelectedCell: PCell;
begin
  if (Worksheet <> nil) then
    with Worksheet do
      Result := FindCell(ActiveCellRow, ActiveCellCol)
  else
    Result := nil;
end;

function TsRichCellEdit.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

function TsRichCellEdit.GetWorksheet: TsWorksheet;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Worksheet
  else
    Result := nil;
end;

procedure TsRichCellEdit.ListenerNotification(AChangedItems: TsNotificationItems;
  AData: Pointer = nil);
var
  cell: PCell;
begin
  if (FWorkbookSource = nil) or (FRefocusing = self) then
    exit;

  if  (lniSelection in AChangedItems) or
     ((lniCell in AChangedItems) and (PCell(AData) = SelectedCell))
  then begin
    if Worksheet.IsMerged(SelectedCell) then
      cell := Worksheet.FindMergeBase(SelectedCell)
    else
      cell := SelectedCell;
    ShowCell(cell);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Standard component notification. Called when the WorkbookSource is deleted.
-------------------------------------------------------------------------------}
procedure TsRichCellEdit.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Removes the link of the RichCellEdit to the WorkbookSource. 
  Required before destruction.
-------------------------------------------------------------------------------}
procedure TsRichCellEdit.RemoveWorkbookSource;
begin
  SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the WorkbookSource
-------------------------------------------------------------------------------}
procedure TsRichCellEdit.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
    FWorkbookSource.AddListener(self);
  Text := '';
  SetLength(FRichTextParams, 0);
  ListenerNotification([lniSelection]);
end;

{@@ ----------------------------------------------------------------------------
  Loads the contents of a cell into the editor.
  Shows the formula if available, but not the calculation result.
  Numbers are displayed in full precision.
  Date and time values are shown in the long formats.
  Text values are formatted as rich text.

  @param  ACell  Pointer to the cell loaded into the cell editor.
-------------------------------------------------------------------------------}
procedure TsRichCellEdit.ShowCell(ACell: PCell);
var
  s: String;
  hideformula: Boolean;
begin
  if (FWorkbookSource <> nil) and (ACell <> nil) then
  begin
    hideformula := Worksheet.IsProtected and (spCells in Worksheet.Protection) and
      (cpHideFormulas in Worksheet.ReadCellProtection(ACell));
    s := Worksheet.ReadFormulaAsString(ACell, true);
    SetLength(FRichTextParams, 0);
    if (s <> '') then begin
      if hideformula then
        s := '(Formula hidden)'
      else
        if s[1] <> '=' then s := '=' + s;
      Lines.Text := s;
    end else
      case ACell^.ContentType of
        cctNumber:
          Lines.Text := FloatToStr(ACell^.NumberValue);
        cctDateTime:
          if ACell^.DateTimeValue < 1.0 then        // Time only
            Lines.Text := FormatDateTime('tt', ACell^.DateTimeValue)
          else
          if frac(ACell^.DateTimeValue) = 0 then    // Date only
            Lines.Text := FormatDateTime('ddddd', ACell^.DateTimevalue)
          else                                      // both
            Lines.Text := FormatDateTime('c', ACell^.DateTimeValue);
        cctUTF8String:
          begin
            Lines.Text := ACell^.UTF8StringValue;
            SetLength(FRichTextParams, Length(ACell^.RichTextParams));
            FRichTextParams := ACell^.RichTextParams;
            ApplyRichParams;
          end;
        else
          Lines.Text := Worksheet.ReadAsText(ACell);
      end;
  end else
    Clear;

  FOldText := Lines.Text;

  ReadOnly := not CanEditCell(ACell);
end;

{@@ ----------------------------------------------------------------------------
  Writes the current edit text to the cell

  @Note  All validation checks already have been performed.
-------------------------------------------------------------------------------}
procedure TsRichCellEdit.WriteToCell;
var
  cell: PCell;
  s: String;
begin
  cell := Worksheet.GetCell(Worksheet.ActiveCellRow, Worksheet.ActiveCellCol);
  if Worksheet.IsMerged(cell) then
    cell := Worksheet.FindMergeBase(cell);
  s := Lines.Text;
  if (s <> '') and (s[1] = '=') then
    Worksheet.WriteFormula(cell, Copy(s, 2, Length(s)), true)
  else
  begin
    if cell^.ContentType = cctUTF8String then
      Worksheet.WriteText(cell, s, ExtractRichTextParams)
    else
      Worksheet.WriteCellValueAsString(cell, s);
  end;
end;
      *)
end.

