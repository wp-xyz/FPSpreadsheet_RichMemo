unit unit2;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, 
  fpsTypes, fpSpreadsheet, fpspreadsheetgrid, fpsRichTextCtrls;

type
  
  { TForm1 }

  TForm1 = class(TForm)
    Label1: TLabel;
    sWorksheetGrid1: TsWorksheetGrid;
    procedure FormCreate(Sender: TObject);
    procedure sWorksheetGrid1Click(Sender: TObject);
  private
    memo: TsCustomCellRichMemo;

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.FormCreate(Sender: TObject);
begin
  sWorksheetGrid1.Worksheet.WriteTextAsHTML(0, 0, 'Lorem <i><b>ipsum</b></i>');
  sWorksheetGrid1.Worksheet.WriteTextAsHTML(1, 1, '1 <font color="red">RED</font>-<font color="green">GREEN</font>-<font color="blue">BLUE</font>.');
  sWorksheetGrid1.Worksheet.WriteTextAsHTML(2, 2, 'H<sub>2</sub>O');
  sWorksheetGrid1.Worksheet.WriteTextAsHTML(3, 3, '10 km<sup>2</sup>');
  sWorksheetGrid1.Worksheet.WriteTextAsHTML(4, 4, 'Test <b>abc</b>');
  sWorksheetGrid1.Worksheet.WriteFont(4, 4, 'Courier New', 12, [], scRed);
  sWorksheetGrid1.Worksheet.WriteText(4, 3, 'abc');
  sWorksheetGrid1.Worksheet.WriteFont(4, 3, 'Courier New', 10, [], scGreen);
  
  memo := TsCustomCellRichMemo.Create(self);
  memo.Parent := self;
  memo.Left := sWorksheetGrid1.Left;
  memo.Top := sWorksheetgrid1.Top + sWorksheetGrid1.Height + 8;
  memo.Height := 100;
  memo.Width := 400;
end;

procedure TForm1.sWorksheetGrid1Click(Sender: TObject);
var
  cell: PCell;
  rtp: TsRichTextParam;
  s: String;
begin
  with sWorksheetGrid1 do
    cell := Worksheet.FindCell(Row-1, Col-1);
  
  s := '';
  if cell <> nil then
    for rtp in cell^.RichTextParams do
      s := Format(s + LineEnding + 'rtp.FirstIndex = %d, rtp.FontIndex = %d', [rtp.FirstIndex, rtp.FontIndex]);
  Label1.Caption := s;
  
  memo.CellToMemo(cell);
end;

end.

