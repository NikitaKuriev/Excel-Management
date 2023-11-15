unit Excel;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, COMObj;

type
  TForm2 = class(TForm)
    ButtonCreate: TButton;
    ButtonClose: TButton;
    ButtonOpenExcel: TButton;
    ButtonOpenSheet1: TButton;
    ButtonOpenSheet2: TButton;
    Edit1: TEdit;
    ButtonAddC3: TButton;
    Edit2: TEdit;
    ButtonAddA2: TButton;
    ButtonAddFormula: TButton;
    ButtonFillCell: TButton;
    ButtonFillColStr: TButton;
    Button—ellParam: TButton;
    ButtonCellColParam: TButton;
    procedure ButtonCreateClick(Sender: TObject);
    procedure ButtonCloseClick(Sender: TObject);
    procedure ButtonOpenExcelClick(Sender: TObject);
    procedure ButtonOpenSheet1Click(Sender: TObject);
    procedure ButtonOpenSheet2Click(Sender: TObject);
    procedure ButtonAddC3Click(Sender: TObject);
    procedure ButtonAddA2Click(Sender: TObject);
    procedure ButtonAddFormulaClick(Sender: TObject);
    procedure ButtonFillCellClick(Sender: TObject);
    procedure ButtonFillColStrClick(Sender: TObject);
    procedure Button—ellParamClick(Sender: TObject);
    procedure ButtonCellColParamClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;
  XLApp,WBook,WSheet:OleVariant;
implementation

{$R *.dfm}

procedure TForm2.ButtonOpenSheet2Click(Sender: TObject);
begin
  XLApp:=CreateOleObject('Excel.Application');
  XLApp.Visible:=True;
  WBook:=XLApp.Workbooks.Open('d:\\Listen.xlsx');
  XLApp.Worksheets[2].Activate;
end;

procedure TForm2.Button—ellParamClick(Sender: TObject);
begin
  XLApp:=CreateOleObject('Excel.Application');
  XLApp.Visible:=True;
  XLApp.Workbooks.Add;
  WSheet:=XLApp.Worksheets[1];
  WSheet.Activate;

  WSheet.Range['B5'].Select;
  XLApp.ActiveCell.Value := 'Text';
  XLApp.ActiveCell.Font.Size := 24;
  XLApp.ActiveCell.Font.FontStyle := 'Bold';
  XLApp.ActiveCell.Font.Name := 'Arial';
  XLApp.ActiveCell.Font.Color := RGB(150,0,50);
end;

procedure TForm2.ButtonAddA2Click(Sender: TObject);
begin
  XLApp:=CreateOleObject('Excel.Application');
  XLApp.Visible:=True;
  XLApp.Workbooks.Add;
  WSheet:=XLApp.Worksheets[1];
  WSheet.Activate;
  WSheet.Cells.Item[2,1].Value:=Edit2.Text;
end;

procedure TForm2.ButtonAddC3Click(Sender: TObject);
begin
  XLApp:=CreateOleObject('Excel.Application');
  XLApp.Visible:=True;
  XLApp.Workbooks.Add;
  WSheet:=XLApp.Worksheets[1];
  WSheet.Activate;
  WSheet.Range['C3'].Value:=Edit1.Text;
end;

procedure TForm2.ButtonAddFormulaClick(Sender: TObject);
begin
  XLApp:=CreateOleObject('Excel.Application');
  XLApp.Visible:=True;
  XLApp.Workbooks.Add;

  XLApp.ActiveSheet.Range['A5'].Formula:= '=SUM(A1,A2)*5';
end;

procedure TForm2.ButtonCellColParamClick(Sender: TObject);
var
  F:OleVariant;
begin
  XLApp:=CreateOleObject('Excel.Application');
  XLApp.Visible:=True;
  XLApp.Workbooks.Add;
  WSheet:=XLApp.Worksheets[1];
  WSheet.Activate;

  F:=WSheet.Range['A:A'].Font;
  F.Name:='Arial';
  F.Size:=28;
  F.FontStyle:='Bold';
  F.Color:=RGB(255,0,0);
end;

procedure TForm2.ButtonCloseClick(Sender: TObject);
begin
 if not VarIsEmpty (XLApp) then
      XLApp.Quit;

end;

procedure TForm2.ButtonCreateClick(Sender: TObject);
begin
  XLApp:=CreateOleObject('Excel.Application');
  XLApp.Visible:=True;
  XLApp.Workbooks.Add;
end;

procedure TForm2.ButtonFillCellClick(Sender: TObject);
begin
  XLApp:=CreateOleObject('Excel.Application');
  XLApp.Visible:=True;
  XLApp.Workbooks.Add;
  WSheet:=XLApp.Workbooks[1];
  WSheet.Activate;

  XLApp.ActiveSheet.Range['C5'].Select;
  XLApp.ActiveCell.Interior.Color:=RGB(0,255,0);
end;

procedure TForm2.ButtonFillColStrClick(Sender: TObject);
begin
  XLApp:=CreateOleObject('Excel.Application');
  XLApp.Visible:=True;
  XLApp.Workbooks.Add;
  WSheet:=XLApp.Worksheets[1];
  WSheet.Activate;
  WSheet.Range['B:B'].Interior.Color:=RGB(0,255,0);
  WSheet.Range['2:2'].Interior.Color:=RGB(0,0,255);
end;

procedure TForm2.ButtonOpenExcelClick(Sender: TObject);
begin
  XLApp:=CreateOleObject('Excel.Application');
  XLApp.Visible:=True;
  WBook:=XLApp.Workbooks.Open('d:\\Listen.xlsx');
end;

procedure TForm2.ButtonOpenSheet1Click(Sender: TObject);
begin
  XLApp:=CreateOleObject('Excel.Application');
  XLApp.Visible:=True;
  WBook:=XLApp.Workbooks.Open('d:\\Listen.xlsx');
  WSheet:=XLApp.Worksheets[2];
  WSheet.Activate;
end;

end.
