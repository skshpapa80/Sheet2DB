// 작성자: 조용호 (skshpapa80@gmail.com)
// 프로그래명 : Xls 파일내용을 DB에 입력하기 위한 프로그램
// 작성일 : 2015-11-11
// 수정일 : 2020-07-18
// 블로그 : https://skshpapa80.github.io/
//
// Xls 파일을 선택하여
// DB 정보 접속정보를 입력하고 테이블 정보와 엑셀컬럼을 매칭하여
// 테이블에 값을 인서트 하는 프로그램
//
//
unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Grids, System.Win.ComObj,
  Data.DB, Data.Win.ADODB;

type
  TfrmMain = class(TForm)
    btnClose: TButton;
    ADOQuery1: TADOQuery;
    PageControl1: TPageControl;
    TabSheet2: TTabSheet;
    btnOpen: TButton;
    StringGrid1: TStringGrid;
    TabSheet1: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit3: TEdit;
    Edit1: TEdit;
    Btn_DBConn: TButton;
    TabSheet4: TTabSheet;
    Label5: TLabel;
    StringGrid2: TStringGrid;
    btnColumn: TButton;
    StringGrid3: TStringGrid;
    StringGrid4: TStringGrid;
    StringGrid5: TStringGrid;
    btnAdd: TButton;
    btnDel: TButton;
    TabSheet3: TTabSheet;
    Label6: TLabel;
    ProgressBar1: TProgressBar;
    btnSave: TButton;
    StringGrid6: TStringGrid;
    btnDBRead: TButton;
    Edit2: TEdit;
    Memo1: TMemo;
    procedure btnCloseClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnOpenClick(Sender: TObject);
    procedure Btn_DBConnClick(Sender: TObject);
    procedure btnColumnClick(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnDelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnDBReadClick(Sender: TObject);
  private
    { Private declarations }
    RowCnt, RowCnt1: Integer;
    procedure LoadXLS(filename : string; stgrid : TStringGrid);
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.btnAddClick(Sender: TObject);
begin
    // 컬럼정보와 엑셀 컬럼 연결
    StringGrid5.Cells[0,RowCnt -1] := StringGrid4.Cells[1,StringGrid4.Row];
    StringGrid5.Cells[1,RowCnt -1] := StringGrid3.Cells[3,StringGrid3.Row];
    StringGrid5.Cells[2,RowCnt -1] := StringGrid4.Cells[0,StringGrid4.Row];
    StringGrid5.Cells[3,RowCnt -1] := StringGrid3.Cells[5,StringGrid3.Row];
    StringGrid5.RowCount := RowCnt;
    Inc(RowCnt);
end;

procedure TfrmMain.btnCloseClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmMain.btnColumnClick(Sender: TObject);
var
    i,j : integer;
begin
    // 선택한 테이블의 컬럼정보 가져오기
    try
        ADOQuery1.SQL.Clear;
        ADOQuery1.SQL.Text := 'sp_columns ''' + StringGrid2.Cells[2,StringGrid2.Row] + '''';
        ADOQuery1.Open;

        i:=0;
        StringGrid3.ColCount := ADOQuery1.Recordset.Fields.Count;
        while not ADOQuery1.Eof do begin
            for j:=0 to ADOQuery1.Recordset.Fields.Count -1 do begin
                StringGrid3.Cells[j,i] := ADOQuery1.Fields.Fields[j].AsString;
            end;
            StringGrid3.RowCount := i+1;
            inc(i);
            ADOQuery1.Next;
        end;
    except

    end;
end;

procedure TfrmMain.btnDBReadClick(Sender: TObject);
var
    sqltxt: String;
    i : integer;
begin
    // 결과 확인
    if StringGrid2.Cells[2,StringGrid2.Row] = '' then Exit;

    sqltxt := 'SELECT * FROM ' + StringGrid2.Cells[2,StringGrid2.Row];
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Text := sqltxt;
    ADOQuery1.Open;

    RowCnt := 2;
    StringGrid6.ColCount := ADOQuery1.Recordset.Fields.Count;
    for i := 0 to StringGrid3.RowCount - 1 do begin

        StringGrid6.Cells[i,0] := StringGrid3.Cells[3,i];

    end;

    while not ADOQuery1.Eof do begin
        for i := 0 to ADOQuery1.Recordset.Fields.Count - 1 do begin
            StringGrid6.Cells[i,RowCnt-1] := ADOQuery1.Fields.Fields[i].AsString;
        end;
        StringGrid6.RowCount := RowCnt;
        inc(RowCnt);
        ADOQuery1.Next;
    end;
    ADOQuery1.Close;

    Edit2.Text := IntToSTr(RowCnt-2);
end;

procedure TfrmMain.btnDelClick(Sender: TObject);
var
    i: Integer;
begin
    // String Grid ROW 삭제
    for i := StringGrid5.Row to StringGrid5.RowCount - 2 do
        StringGrid5.Rows[i].Assign(StringGrid5.Rows[i + 1]);
    StringGrid5.RowCount := StringGrid5.RowCount - 1;
    Dec(RowCnt);
end;

procedure TfrmMain.btnOpenClick(Sender: TObject);
var
    i : integer;
begin
    // 엑셀파일 열기
    with TOpenDialog.Create(Self) do begin
        Title := '엑셀 파일 열기';
        Filter := '엑셀파일(*.xlsx,*.xls)|*.xlsx;*.xls';
        if Execute then begin
            LoadXLS(FileName, StringGrid1);

            for i :=0 to StringGrid1.ColCount - 1 do begin
                StringGrid4.Cells[0,i] := IntToStr(i);
                StringGrid4.Cells[1,i] := StringGrid1.Cells[i,0];
                StringGrid4.RowCount := i+1;
            end;
        end;
    end;
end;

procedure TfrmMain.btnSaveClick(Sender: TObject);
var
    sqltxt: string;
    sqltxt1: string;
    i, j, cnt: integer;
    tmp: String;
begin
    // 데이터 저장
    if StringGrid5.Cells[0,1] = '' then begin
        ShowMessage('매칭된 컬럼이 없습니다!');
        Exit;
    end;

    sqltxt := 'insert into ' + StringGrid2.Cells[2,StringGrid2.Row] + ' (';
    for j:= 1 to StringGrid5.RowCount -1 do begin
        sqltxt := sqltxt + StringGrid5.Cells[1,j];
        if j < StringGrid5.RowCount - 1 then
            sqltxt := sqltxt + ',';
    end;
    sqltxt := sqltxt + ') values(';


    ProgressBar1.Min := 1;
    ProgressBar1.Max := StringGrid1.RowCount;
    ProgressBar1.Position := 1;
    ProgressBar1.Update;
    cnt := 0;
    for i:= 1 to StringGrid1.RowCount - 1 do begin
        if StringGrid1.Cells[0,i] = '' then Continue;
        sqltxt1 := sqltxt;
        for j := 1 to StringGrid5.RowCount - 1 do begin
            tmp := StringGrid1.Cells[0,i];
            if StringGrid5.Cells[3,j] = 'numeric' then begin
                if StringGrid1.Cells[StrToInt(StringGrid5.Cells[2,j]),i] = '' then
                    sqltxt1 := sqltxt1 + '''0'''
                else
                    sqltxt1 := sqltxt1 + ''''+StringGrid1.Cells[StrToInt(StringGrid5.Cells[2,j]),i]+'''';
            end
            else begin
                sqltxt1 := sqltxt1 + ''''
                          +StringReplace(StringGrid1.Cells[StrToInt(StringGrid5.Cells[2,j]),i],'''','''''',[rfReplaceAll])+'''';
            end;
            if j < StringGrid5.RowCount - 1 then
                sqltxt1 := sqltxt1 + ',';
        end;
        sqltxt1 := sqltxt1 + ')';
        //ShowMessage(sqltxt1);

        try
            ADOQuery1.SQL.Clear;
            ADOQuery1.SQL.Text := sqltxt1;
            ADOQuery1.ExecSQL;
        except
            Inc(RowCnt1);

            ADOQuery1.Close;
            ADOQuery1.Open;

            Memo1.Text := sqltxt1;
            //showmessage(tmp);
            Exit;
        end;
        inc(cnt);
        ProgressBar1.Position := ProgressBar1.Position+1;
        Application.ProcessMessages;
    end;

    ShowMessage('DB입력이 ' + IntToStr(Cnt) + '건 완료되었습니다.');
end;

procedure TfrmMain.Btn_DBConnClick(Sender: TObject);
var
    i,j : integer;
begin
    // DB 접속 하기
    //ADOQuery1.ConnectionString := Edit_Conn.Text;
    ADOQuery1.ConnectionString := ' Provider=SQLNCLI11.1;Password=' + edit4.text
                                + ';Persist Security Info=True;User ID=' + edit5.Text
                                + ';Initial Catalog=' + edit3.Text + ';Data Source=' + edit1.Text + ' ';

    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Text := ' sp_tables ';
    ADOQuery1.Open;

    i:=0;
    // 테이블 정보 불러오기
    while not ADOQuery1.Eof do begin
        for j := 0 to ADOQuery1.Recordset.Fields.Count -1 do begin
            StringGrid2.Cells[j,i] := ADOQuery1.Fields.Fields[j].AsString;
        end;
        StringGrid2.RowCount := i+1;
        inc(i);
        ADOQuery1.Next;
    end;

    PageControl1.TabIndex := 2;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
    PageControl1.TabIndex := 0;

    StringGrid5.Cells[0,0] := '엑셀네임';
    StringGrid5.Cells[1,0] := '필드명';
    RowCnt := 2;
    RowCnt1 := 2;
end;

procedure TfrmMain.LoadXLS(filename: string; stgrid: TStringGrid);
var
    oXL, oWK, oSheet: Variant;
    i, j, row : Integer;
begin
    try
        // Excel Automation
        oXL := CreateOleObject('Excel.Application');
        oXL.Visible := False;
        oXL.DisplayAlerts := False;
        oXL.WorkBooks.Open(filename, 0, true); // 읽기 전용으로 읽기

        oWK := oXL.WorkBooks.Item[1];
        oSheet := oWK.ActiveSheet;

        row := 0;
        stgrid.RowCount := StrToInt(oSheet.UsedRange.Rows.count);
        stgrid.ColCount := StrToInt(oSheet.UsedRange.Columns.count);

        for i := 1 to StrToInt(oSheet.UsedRange.Rows.count) do begin
            for j := 1 to StrToInt(oSheet.UsedRange.Columns.count) do begin
                stgrid.Cells[j - 1, row] := VarToStr(oSheet.Cells[i,j]);
            end;
            row := row + 1;
        end;

        oXL.WorkBooks.Close;
        oXL.Quit;
        oXL := unassigned;
    except
	    Exit;
    end;
end;

end.
