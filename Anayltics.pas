
unit Unit3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, ADODB, Grids, DBGrids, ComCtrls, TeEngine, ExtCtrls,
  TeeProcs, Chart, Series, DBCtrls, DBCGrids, DBTables, WideStrings, FMTBcd,
  SqlExpr, DBXOracle;

type
  TForm3 = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    dsOrder: TDataSource;
    TabSheet2: TTabSheet;
    DBGrid1: TDBGrid;
    ADOQrder: TADOQuery;
    Button1: TButton;
    EdtOrder: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Button2: TButton;
    DBGrid2: TDBGrid;
    ADOEmp: TADOQuery;
    dsEmp: TDataSource;
    BtnExpOrd: TButton;
    Button4: TButton;
    DateTimePickerEmpBeg: TDateTimePicker;
    TabSheet3: TTabSheet;
    ADOType: TADOQuery;
    Chart1: TChart;
    Series1: TPieSeries;
    Button5: TButton;
    DateTimePickerTypeB: TDateTimePicker;
    ADOTypeORDERTYPE: TStringField;
    ADOTypeCOUNTORDERTYPE: TFloatField;
    Label3: TLabel;
    ADOOrdbyEmp: TADOQuery;
    ADOOrdbyEmpEMPLOYEE: TStringField;
    ADOOrdbyEmpCOUNTORDERNO: TFloatField;
    ADOTypByEmp: TADOQuery;
    ADOTypByEmpORDERTYPE: TStringField;
    ADOTypByEmpCOUNTORDERTYPE: TFloatField;
    TabSheet6: TTabSheet;
    DBGrid3: TDBGrid;
    dsOrdByDts: TDataSource;
    ADOOrdByDt: TADOQuery;
    TabSheet7: TTabSheet;
    Chart4: TChart;
    ADOQC: TADOQuery;
    BtnQC: TButton;
    DateTimePickerQCb: TDateTimePicker;
    Label8: TLabel;
    ADOQCEMPLOYEE: TStringField;
    ADOQCCOUNTORDERNO: TFloatField;
    Series4: TPieSeries;
    TabSheet8: TTabSheet;
    Chart5: TChart;
    ADOStation: TADOQuery;
    Series5: TPieSeries;
    BtnStation: TButton;
    DateTimeStatE: TDateTimePicker;
    ADOStationSTATION: TStringField;
    ADOStationCOUNTSTATION: TFloatField;
    Label9: TLabel;
    Label4: TLabel;
    TabSheet4: TTabSheet;
    ADODupChk: TADOQuery;
    ADODspDup: TADOQuery;
    DateTimePickerDupBeg: TDateTimePicker;
    BtnDup: TButton;
    dsDup: TDataSource;
    ADODupChkORDERNO: TFMTBCDField;
    ADODupChkCNT: TFloatField;
    Label5: TLabel;
    ADOEmpStation: TADOQuery;
    ADODspDupORDERNO: TFMTBCDField;
    ADODspDupEMPLOYEE: TStringField;
    ADODspDupSTATION: TStringField;
    ADODspDupDTS: TStringField;
    Memo1: TMemo;
    Label12: TLabel;
    ADODisOrdByDts: TADOQuery;
    DBGrid4: TDBGrid;
    DsEmpStation: TDataSource;
    ADOEmpStationSTATION: TStringField;
    ADOEmpStationCOUNTSTATION: TFloatField;
    QryName: TADOQuery;
    QryNamefstnam: TWideMemoField;
    EdtEmpName: TEdit;
    ADOFullName: TADOQuery;
    Chart2: TChart;
    Series2: TPieSeries;
    ADOEmpTyp: TADOQuery;
    ADOCntType: TADOQuery;
    ADOEmpTypORDERNO: TFMTBCDField;
    ADOCntTypeORDERTYPE: TStringField;
    dsDlyryType: TDataSource;
    ADODlryType: TADOQuery;
    Memo2: TMemo;
    ADODlryTypeDLVRY_TYPE: TStringField;
    ADODlryTypeCOUNTDLVRY_TYPE: TFloatField;
    EdtRecords: TEdit;
    Label6: TLabel;
    Label11: TLabel;
    TabSheet5: TTabSheet;
    Chart3: TChart;
    BtnEmpOrds: TButton;
    ADOAllEmps: TADOQuery;
    ADOAllEmpsEMPLOYEE: TStringField;
    ADOAllEmpsCOUNTEMPLOYEE: TFloatField;
    Series3: TBarSeries;
    Memo3: TMemo;
    dsAmyTm: TDataSource;
    ADOAmyTm: TADOQuery;
    ADOAmyTmLASTNAM: TWideMemoField;
    ADOAmyTmFSTNAM: TWideMemoField;
    ADOAmyTmemp_no: TWideMemoField;
    ADOFullNamefstnam: TWideMemoField;
    ADOFullNamelastnam: TWideMemoField;
    ADOFullNameemp_no: TWideMemoField;
    ADOGetName: TADOQuery;
    ADOGetNamefstnam: TWideMemoField;
    ADOGetNamelastnam: TWideMemoField;
    ComboBox1: TComboBox;
    ADOQrderEMPLOYEE: TStringField;
    ADOQrderORDERNO: TFMTBCDField;
    ADOQrderSTATION: TStringField;
    ADOQrderORDERTYPE: TStringField;
    ADOQrderDTS: TStringField;
    ADOQrderPKG_CNT: TStringField;
    ADOQrderDLVRY_TYPE: TStringField;
    ADOQrderCMNTS: TStringField;
    ADOName: TADOQuery;
    ADONamelastnam: TWideMemoField;
    ADOQrderLastnm: TStringField;
    ADOOrdByDtLastNm: TStringField;
    ADOOrdByDtORDERNO: TFMTBCDField;
    ADOOrdByDtSTATION: TStringField;
    ADOOrdByDtEMPLOYEE: TStringField;
    ADOOrdByDtORDERTYPE: TStringField;
    ADOOrdByDtDTS: TStringField;
    ADOOrdByDtPKG_CNT: TStringField;
    ADOOrdByDtDLVRY_TYPE: TStringField;
    ADOOrdByDtCMNTS: TStringField;
    TabSheet9: TTabSheet;
    ADOInComplete: TADOQuery;
    DsInComplete: TDataSource;
    ADOInComplete1: TADOQuery;
    ADOInCompleteORDERNO: TFMTBCDField;
    Memo4: TMemo;
    ADOTracking: TADOQuery;
    ADOTrackingshipmentid: TStringField;
    ADOTrackingtouchdate: TDateTimeField;
    ADOOrdByDtShipDts: TStringField;
    ADONamefstnam: TWideMemoField;
    ADOQrderShipDts: TStringField;
    Chart6: TChart;
    ADOChrtInClmp: TADOQuery;
    ADOChrtInClmpORDERTYPE: TStringField;
    ADOChrtInClmpCOUNTORDERTYPE: TFloatField;
    Series6: TPieSeries;
    Series7: TPieSeries;
    ADOOrdSeq: TADOQuery;
    Chart7: TChart;
    Series8: TPieSeries;
    Series9: TBarSeries;
    Series10: TBarSeries;
    Panel1: TPanel;
    DateTimePickerTQCb: TDateTimePicker;
    BtnInCompl: TButton;
    Label13: TLabel;
    EdtInRecords: TEdit;
    Button3: TButton;
    Panel2: TPanel;
    DateTimePickerAllEmpsB: TDateTimePicker;
    ADOInComplete1ORDERNO: TFMTBCDField;
    ADOInComplete1STATION: TStringField;
    ADOInComplete1EMPLOYEE: TStringField;
    ADOInComplete1ORDERTYPE: TStringField;
    ADOInComplete1DTS: TStringField;
    ADOInComplete1PKG_CNT: TStringField;
    ADOInComplete1DLVRY_TYPE: TStringField;
    ADOInComplete1CMNTS: TStringField;
    ADOInComplete1NDTS: TDateTimeField;
    ADOInComplete1LastNm: TStringField;
    ADOSeqTQ: TADOQuery;
    ADOSeqTQORDERNO: TFMTBCDField;
    Button6: TButton;
    ADOOrdSeqORDERTYPE: TStringField;
    Series11: TPieSeries;
    SQLspTQ: TSQLStoredProc;
    Label14: TLabel;
    DateTimePickerEmpEnd: TDateTimePicker;
    spTrAndAssgn: TSQLStoredProc;
    DateTimePickerTypeE: TDateTimePicker;
    DateTimePickerQCe: TDateTimePicker;
    DateTimePickerTQCe: TDateTimePicker;
    spTtoQCdr: TSQLStoredProc;
    DateTimePickerAllEmpsE: TDateTimePicker;
    DateTimeStatB: TDateTimePicker;
    Label7: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    SQLConnection2: TSQLConnection;
    ADOTriage: TADOQuery;
    ADOTriageORDERTYPE: TStringField;
    ADOTriageCOUNTORDERTYPE: TFloatField;
    BtnExport: TButton;
    DateTimePickerDupEnd: TDateTimePicker;
    Label19: TLabel;
    Errors: TTabSheet;
    DBGrid5: TDBGrid;
    ADOErrsByDate: TADOQuery;
    dsErrors: TDataSource;
    ADOErrsByDateORDERNO: TFMTBCDField;
    ADOErrsByDateEMPLOYEE: TStringField;
    ADOErrsByDateERROR_TYPE: TStringField;
    ADOErrsByDateERROR_NO: TStringField;
    ADOErrsByDateERROR_DSC: TStringField;
    ADOErrsByDateNDTS: TDateTimeField;
    ADOErrsByDateLastNm: TStringField;
    ADOErrsByDateDesc: TStringField;
    GetDesc: TADOQuery;
    Panel3: TPanel;
    Label20: TLabel;
    DateTimeErrBeg: TDateTimePicker;
    Label21: TLabel;
    DateTimeErrEnd: TDateTimePicker;
    BtnError: TButton;
    Button10: TButton;
    ADOErrsByDateTDTS: TStringField;
    Panel4: TPanel;
    Label15: TLabel;
    DateTimePickerBeg: TDateTimePicker;
    Label16: TLabel;
    DateTimePickerEnd: TDateTimePicker;
    Button8: TButton;
    Button9: TButton;
    Label10: TLabel;
    EdtTotal: TEdit;
    GetDescERROR_DSC: TStringField;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure BtnExpOrdClick(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure BtnQCClick(Sender: TObject);
    procedure BtnStationClick(Sender: TObject);
    procedure BtnDupClick(Sender: TObject);
    procedure DBGrid2TitleClick(Column: TColumn);
    procedure DBGrid3TitleClick(Column: TColumn);
    procedure BtnEmpOrdsClick(Sender: TObject);
    procedure ComboBox1Click(Sender: TObject);
    procedure ADOQrderCalcFields(DataSet: TDataSet);
    procedure BtnInComplClick(Sender: TObject);
    procedure ADOInComplete1CalcFields(DataSet: TDataSet);
    procedure DBGrid3CellClick(Column: TColumn);
    procedure Button3Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Memo4KeyPress(Sender: TObject; var Key: Char);
    procedure DateTimePickerBegChange(Sender: TObject);
    procedure DateTimePickerEmpBegChange(Sender: TObject);
    procedure DateTimePickerTypeBChange(Sender: TObject);
    procedure DateTimePickerQCbChange(Sender: TObject);
    procedure DateTimeStatBChange(Sender: TObject);
    procedure DateTimePickerAllEmpsBChange(Sender: TObject);
    procedure DateTimePickerTQCbChange(Sender: TObject);
    procedure BtnExportClick(Sender: TObject);
    procedure BtnErrorClick(Sender: TObject);
    procedure DBGrid5TitleClick(Column: TColumn);
    procedure Button10Click(Sender: TObject);
    procedure DBGrid5CellClick(Column: TColumn);


  private

  public
    var sort: String;
    var sortbydts: String;
    var CbName,lsst,frst: String;
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

procedure TForm3.Button10Click(Sender: TObject);
var
i: integer;
strT: string;
slst: TStringList;
begin
   screen.Cursor := crHourGlass;
   slst:= TStringList.Create;
try
with dsErrors.DataSet do
  begin
        First;
    while not Eof do
    begin
      strT:= '"'+Fields[0].AsString+'"';
      for i:= 1 to FieldCount-1 do
        strT:= strT+',"'+Fields[i].AsString+'"';
      slst.Add(strT);
      Next;
    end;
        First;
  end;
  slst.SaveToFile('C:\Data\ByDate.csv');
Finally
  slst.Free;
end;
screen.Cursor := crDefault;
end;

procedure TForm3.Button1Click(Sender: TObject);
begin

     with ADOQrder do
        begin
          ADOQrder.Close;
          ADOQrder.Parameters.ParamByName('orderno').Value := EdtOrder.Text;
          ADOQrder.Open;
        end;
end;



procedure TForm3.Button2Click(Sender: TObject);
var dts,dtsB,dtsE, Type1: String;
var spDtsB,spDtsE: TDateTime;
var i,j,k,l,m,n,o,p: Integer;
var sort1: String;
var EmpNumber: String;
var iv,ent,inn,wc,pu,rs: String;


begin
i := 0;
j := 0;
k := 0;
l := 0;
m := 0;
n := 0;
o := 0;
p := 0;

sort1 := sort;
screen.Cursor := crHourGlass;

spDtsB := DateTimePickerEmpBeg.DateTime;
dtsB := FormatDateTime('dd-MMM-yy',DateTimePickerEmpBeg.DateTime);
dtsB := UpperCase(dtsB);

spDtsE := DateTimePickerEmpEnd.DateTime;
dtsE := FormatDateTime('dd-MMM-yy',DateTimePickerEmpEnd.DateTime);
dtsE := UpperCase(dtsE);

  with ADOFullName do
      begin
        ADOFullName.Close;
        ADOFullName.Parameters.ParamByName('%lastname%').Value := lsst;
        ADOFullName.Parameters.ParamByName('%firstname%').Value := frst;
        ADOFullName.Open;
      end;

     EmpNumber := copy(ADOFullNameemp_no.AsString,1,6);
     EdtEmpName.Text := ADOFullNamefstnam.AsString + ' ' + ADOFullNamelastnam.AsString;

     with ADOEmp do
        begin
          Close;
          sql.Clear;
          sql.Add('Select * from mcca.hc_whs');
          sql.Add('where employee = ''' + EmpNumber + '''');
          sql.Add('and ndts > ''' + dtsB + '''');
          sql.Add('and ndts < ''' + dtsE + '''');

          if  sort1 = '0' then
           sql.Add('order by orderno')
          else if sort1 = '1' then
          sql.Add('order by ordertype')
          else if sort1 = '2' then
          sql.Add('order by station')
          else if sort1 = '3' then
          sql.Add('order by dlvry_type')
          else if sort1 = '4' then
          sql.Add('order by pkg_cnt')
          else if sort1 = '5' then
          sql.Add('order by ndts')
          else
          sql.Add('order by orderno');
          Open;
          if adoemp.RecordCount = 0 then
              Showmessage('No Records Found for '+ EdtEmpName.Text);
           EdtRecords.Text := IntToStr(adoemp.RecordCount);
        end;

  with ADOEmpStation do
        begin
          ADOEmpStation.Close;
          ADOEmpStation.Parameters.ParamByName('employee').Value := EmpNumber;
          ADOEmpStation.Parameters.ParamByName('date3').Value := dtsB;
          ADOEmpStation.Parameters.ParamByName('date4').Value := dtsE;
          ADOEmpStation.Open;
        end;


  with spTrAndAssgn do
  begin
      Params.ParamByName('p_Bdate').value := spDtsB;
      Params.ParamByName('p_Edate').value := spDtsE;
      Params.ParamByName('p_emp').value := EmpNumber;
      ExecProc;
      iv := spTrAndAssgn.Params.ParamByName('IV_RTN').value;
      ent := spTrAndAssgn.Params.ParamByName('ENT_RTN').value;
      inn := spTrAndAssgn.Params.ParamByName('INN_RTN').value;
      wc := spTrAndAssgn.Params.ParamByName('WC_RTN').value;
      pu := spTrAndAssgn.Params.ParamByName('PU_RTN').value;
      rs := spTrAndAssgn.Params.ParamByName('RS_RTN').value;
    end;

      Series2.Clear;
      Chart2.Visible := true;
      if adoemp.RecordCount = 0 then
              Chart2.Visible := false;
      with Series2 do
      begin
        Add( strtoint(iv), 'IV' , clRed );
        Add( strtoint(ent), 'Entral' , clBlue );
        Add( strtoint(inn), 'Incontinence' , clGreen );
        Add( strtoint(wc), 'Will Call', clPurple );
//        Add( strtoint(pu), 'Pick Up' , clYellow );
//        Add( strtoint(rs), 'ReSupply' , clNavy );
      end;
      screen.Cursor := crDefault;
end;


procedure TForm3.Button6Click(Sender: TObject);
var dtsB,dtsE, Namet: String;
begin
   dtsB := FormatDateTime('dd-MMM-yy',DateTimePickerTQCb.DateTime);
    dtsB := UpperCase(dtsB);

    dtsE := FormatDateTime('dd-MMM-yy',DateTimePickerTQCe.DateTime);
    dtsE := UpperCase(dtsE);

  with ADOInComplete do
  begin
    ADOInComplete.Close;
    ADOInComplete.Parameters.ParamByName('dtsb').Value := dtsB;
    ADOInComplete.Parameters.ParamByName('dtse').Value := dtsE;
    ADOInComplete.Open;
end;
  EdtInRecords.Text := IntToStr(ADOInComplete.RecordCount);
end;


procedure TForm3.BtnErrorClick(Sender: TObject);
var dts, dtsB,dtsE,sortdts: String;
begin
    dtsB := FormatDateTime('dd-MMM-yy',DateTimeErrBeg.DateTime);
  dtsB := UpperCase(dtsB);

  dtsE := FormatDateTime('dd-MMM-yy',DateTimeErrEnd.DateTime);
  dtsE := UpperCase(dtsE);

  sortdts := sortbydts;

  screen.Cursor := crHourGlass;
  with ADOErrsByDate do
        begin
          Close;
          sql.Clear;
          sql.Add('Select * from mcca.hc_whs_errors');
          sql.Add('where ndts > ''' + dtsB + '''');
          sql.Add('and ndts < ''' + dtsE + '''');
          if  sortdts = '0' then
          sql.Add('order by ndts')
          else if  sortdts = '1' then
          sql.Add('order by employee')
          else if sortdts = '2' then
          sql.Add('order by orderno')
          else if sortdts = '3' then
          sql.Add('order by error_type')
          else if sortdts = '4' then
          sql.Add('order by error_no')
          else if sortdts = '5' then
          sql.Add('order by error_dsc')
          else if sortdts = '6' then
          sql.Add('order by desc')
          else
          sql.Add('order by orderno');
          Open;
          while not ADOErrsByDate.Eof do
          begin
           TStringGrid(DBGrid5).DefaultRowHeight := 25;
           next;
          end;
        end;

   screen.Cursor := crDefault;
end;

procedure TForm3.Button3Click(Sender: TObject);
var dtsB,dtsE, Namet: String;
var i,frst: integer;
begin
    i := 0;
    frst := 5;
    dtsB := FormatDateTime('dd-MMM-yy',DateTimePickerTQCb.DateTime);
    dtsB := UpperCase(dtsB);

    dtsE := FormatDateTime('dd-MMM-yy',DateTimePickerTQCe.DateTime);
    dtsE := UpperCase(dtsE);


   with ADOInComplete do
  begin
    ADOInComplete.Close;
    ADOInComplete.Parameters.ParamByName('dtsb').Value := dtsB;
    ADOInComplete.Parameters.ParamByName('dtse').Value := dtsE;
    ADOInComplete.Open;
    ADOInComplete.First;
    while not eof do
    begin
      with ADOInComplete1 do
      begin
        ADOInComplete1.Close;
        ADOInComplete1.Parameters.ParamByName('order').Value := ADOInCompleteORDERNO.AsString;
        ADOInComplete1.Open;
      end;

        i := i + 1;
       if (i > frst) then
       begin
       if MessageDlg('Continue',
        mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        exit
        else
        frst := frst + 5;
       end;

      Memo4.Lines.Add(ADOInComplete1ORDERNO.AsString +
     '    ' +  ADOInComplete1LastNm.AsString +
      '    ' +  ADOInComplete1STATION.AsString +
      '    ' +  ADOInComplete1ORDERTYPE.AsString +
      '    ' +  ADOInComplete1DTS.AsString);
    next;
  end;
  end;

end;


procedure TForm3.BtnInComplClick(Sender: TObject);
var dtsB,dtsE, Namet: String;
var spDtsB, spDtsE: TDateTime;
var s,t,Integer1, ordcnt: Integer;
var typ: String;
var type1:String;
var amount: Integer;
var iv,ent,inn,wc,pu,rs: String;
var i,j,k,l,m,n,o,p: Integer;
begin
screen.Cursor := crHourGlass;
Series9.Clear;
Series10.Clear;

s := 0;
t := 0;

spDtsB := DateTimePickerTQCb.DateTime;
spDtsE := DateTimePickerTQCe.DateTime;


dtsB := FormatDateTime('dd-MMM-yy',DateTimePickerTQCb.DateTime);
dtsB := UpperCase(dtsB);

dtsE := FormatDateTime('dd-MMM-yy',DateTimePickerTQCe.DateTime);
dtsE := UpperCase(dtsE);

Memo4.Clear;

with ADOType do
begin
  ADOType.Close;
  ADOType.Parameters.ParamByName('dtsb').Value := dtsB;
  ADOType.Parameters.ParamByName('dtse').Value := dtsE;
  ADOType.Open;
  First;
  while not eof do
  begin
    with Series9 do
    begin
          s := s + 668000;
          if ADOTypeCOUNTORDERTYPE.AsInteger < 2 then
          next
          else
          begin
            typ := ADOTypeORDERTYPE.AsString;
            ordcnt := ADOTypeCOUNTORDERTYPE.AsInteger;
            Add(ordcnt, typ, clRed + s );
            next;
          end;
    end;
  end;
end;

with spTtoQCdr do
  begin
      Params.ParamByName('p_Bdate').value := spDtsB;
      Params.ParamByName('p_Edate').value := spDtsE;
      ExecProc;
      iv := spTtoQCdr.Params.ParamByName('IV_RTN').value;
      ent := spTtoQCdr.Params.ParamByName('ENT_RTN').value;
      inn := spTtoQCdr.Params.ParamByName('INN_RTN').value;
      wc := spTtoQCdr.Params.ParamByName('WC_RTN').value;
      pu := spTtoQCdr.Params.ParamByName('PU_RTN').value;
      rs := spTtoQCdr.Params.ParamByName('RS_RTN').value;
    end;

      with Series10 do
      begin
        Add( strtoint(iv), 'IV' , clRed );
        Add( strtoint(ent), 'Entral' , clBlue );
        Add( strtoint(inn), 'Incontinence' , clGreen );
//        Add( strtoint(wc), 'Will Call', clPurple );
//        Add( strtoint(pu), 'Pick Up' , clYellow );
        Add( strtoint(rs), 'ReSupply' , clYellow );
      end;
      screen.Cursor := crDefault;
end;





procedure TForm3.BtnStationClick(Sender: TObject);
var dtsT, type1:String;
var dtsB, dtsE: String;
var station, emp2: String;
var ordcnt,i,statcnt: Integer;
begin
  Series5.Clear;
   i := 0;
//   clRed := $0000FF
//   clBlue := $FF0000
//   clGreen := $008000


 dtsB := FormatDateTime('dd-MMM-yy',DateTimeStatB.DateTime);
 dtsB := UpperCase(dtsB);

 dtsE := FormatDateTime('dd-MMM-yy',DateTimeStatE.DateTime);
 dtsE := UpperCase(dtsE);

with ADOStation do
begin
  ADOStation.Close;
  ADOStation.Parameters.ParamByName('dtsb').Value := dtsB;
  ADOStation.Parameters.ParamByName('dtse').Value := dtsE;
  ADOStation.Open;
  First;
  while not eof do
  begin
    with Series5 do
    begin
       i := i + 1500000;
      station := ADOStationSTATION.AsString;
      statcnt := ADOStationCOUNTSTATION.AsInteger;
       Add( statcnt, station, clGreen + i );
     next;
  end;
end;
end;
end;


procedure TForm3.BtnQCClick(Sender: TObject);
var dtsB, dtsE, type1:String;
var emp, emp1, emp2: String;
var ordcnt,i: Integer;
begin
  Series4.Clear;
//  Chart4.Series[4].Marks.Visible := False;

  screen.Cursor := crHourGlass;
  Memo2.Clear;
   i := 0;
  dtsB := FormatDateTime('dd-MMM-yy',DateTimePickerQCb.DateTime);
  dtsB := UpperCase(dtsB);

  dtsE := FormatDateTime('dd-MMM-yy',DateTimePickerQCe.DateTime);
  dtsE := UpperCase(dtsE);

 with ADODlryType do
 begin
  ADODlryType.Close;
  ADODlryType.Parameters.ParamByName('dtsb').Value := dtsB;
  ADODlryType.Parameters.ParamByName('dtse').Value := dtsE;
  ADODlryType.Open;
  while not eof do
  begin
    Memo2.Font.Size := 8;
    if ADODlryTypeDLVRY_TYPE.AsString = 'U' then
    Memo2.Lines.Add('UPS ' + ADODlryTypeCOUNTDLVRY_TYPE.AsString);
    if ADODlryTypeDLVRY_TYPE.AsString = 'D' then
    Memo2.Lines.Add('DELIVERY ' + ADODlryTypeCOUNTDLVRY_TYPE.AsString);
    if ADODlryTypeDLVRY_TYPE.AsString = 'C' then
    Memo2.Lines.Add('COURIER ' + ADODlryTypeCOUNTDLVRY_TYPE.AsString);
    next;
  end;
  end;

with ADOQC do
begin
  ADOQC.Close;
  ADOQC.Parameters.ParamByName('dtsb').Value := dtsB;
  ADOQC.Parameters.ParamByName('dtse').Value := dtsE;
  ADOQC.Open;
  First;
  while not eof do
  begin
    with QryName do
      begin
        QryName.Close;
        QryName.Parameters.ParamByName('empnum').Value := ADOQCEMPLOYEE.AsString;
        QryName.Open;
      end;
      with Series4 do
      begin
//        i := i + 668000;
         i := i + 96800;
        emp := QryNamefstnam.AsString;
        if length(emp) < 2 then
        next
      else
      begin
        ordcnt := ADOQCCOUNTORDERNO.AsInteger;
        Add( ordcnt, emp, clRed + i );
        next;
      end;
  end;
end;
end;
screen.Cursor := crDefault;
end;

procedure TForm3.ADOInComplete1CalcFields(DataSet: TDataSet);
begin
  with ADOName do
     begin
       Close;
       ADOName.Parameters.ParamByName('empno').Value := DataSet['employee'];
       Open;
     end;
     DataSet['LastNm'] := ADONamelastnam.AsString;
end;

procedure TForm3.ADOQrderCalcFields(DataSet: TDataSet);
begin
    with ADOName do
     begin
       Close;
       ADOName.Parameters.ParamByName('empno').Value := DataSet['employee'];
       Open;
     end;
     DataSet['LastNm'] := ADONamelastnam.AsString;

     with ADOTracking do
     begin
       Close;
       ADOTracking.Parameters.ParamByName('orderno').Value := DataSet['ORDERNO'];
       Open;
     end;
     DataSet['ShipDts'] := ADOTrackingtouchdate.AsString;

end;



procedure TForm3.BtnDupClick(Sender: TObject);

var beg, type1:String;
var bDts, eDts: String;
begin

bDts := FormatDateTime('dd-MMM-yy',DateTimePickerDupBeg.DateTime);
bDts := UpperCase(bDts);

eDts := FormatDateTime('dd-MMM-yy',DateTimePickerDupEnd.DateTime);
eDts := UpperCase(eDts);

  with ADODupChk do
     begin
        Close;
        ADODupChk.Parameters.ParamByName('dtsbeg').Value := bDts;
        ADODupChk.Parameters.ParamByName('dtsend').Value := eDts;
        Open;
        First;
        Memo1.Clear;
        Memo1.Font.Size := 12;
        Memo1.Lines.Add('ORDER   EMPLOYEE   STATION   DATE');
        Memo1.Lines.Add('');
        Memo1.Lines.Add('');
       while not eof do
        begin
          with ADODspDup do
          begin
            Close;
            Parameters.ParamByName('orderno').Value := ADODupChkORDERNO.AsString;
            Open;
            while not eof do
            begin
              Memo1.Font.Size := 10;
              Memo1.Lines.Add(ADODspDupORDERNO.AsString + '       '
              + ADODspDupEMPLOYEE.AsString + '      '
              + ADODspDupSTATION.AsString+ '      '
              + ADODspDupDTS.AsString);
              next;
            end;
            ADODupChk.next;
          end;
        end;
     end;
end;


procedure TForm3.BtnEmpOrdsClick(Sender: TObject);
var dtsB, dtsE, emp, empl, empname: String;
var ordcnt,tcnt,acnt,qcnt,scnt, i: Integer;
var stcnt,st,tr,asn,qc,sg: String;
begin
 i := 0;
 tr := 'T=0';
 asn := 'A=0';
 qc := 'Q=0';
 sg := 'S=0';
 empname := '';


 dtsB := FormatDateTime('dd-MMM-yy',DateTimePickerAllEmpsB.DateTime);
 dtsB := UpperCase(dtsB);

 dtsE := FormatDateTime('dd-MMM-yy',DateTimePickerAllEmpsE.DateTime);
 dtsE := UpperCase(dtsE);



 series11.Clear;
 Memo3.Clear;

 screen.Cursor := crHourGlass;

 with ADOAllEmps do
  begin
  ADOAllEmps.Close;
  ADOAllEmps.Parameters.ParamByName('dts5').Value := dtsB;
  ADOAllEmps.Parameters.ParamByName('dts6').Value := dtsE;
  ADOAllEmps.Open;
  while not eof do
  begin
    Memo3.Lines.Add('');
    with ADOGetName do
    begin
      ADOGetName.Close;
      ADOGetName.Parameters.ParamByName('empnum').Value := ADOAllEmpsEMPLOYEE.AsString;
      ADOGetName.Open;
    end;

    with ADOEmpStation do
    begin
      ADOEmpStation.Close;
      ADOEmpStation.Parameters.ParamByName('employee').Value := ADOAllEmpsEMPLOYEE.AsString;
      ADOEmpStation.Parameters.ParamByName('date3').Value := dtsB;
      ADOEmpStation.Parameters.ParamByName('date4').Value := dtsE;
      ADOEmpStation.Open;
      while not eof do
      begin
        if (length(ADOGetNamelastnam.AsString) < 2) then
        next
        else
        begin
          Memo3.Font.Size := 10;
          if not (empname = ADOGetNamelastnam.AsString)  then
          begin
            empname := ADOGetNamelastnam.AsString;
            Memo3.Lines.Add(ADOGetNamelastnam.AsString);
            Memo3.Lines.Add(ADOEmpStationSTATION.AsString + '  ' + ADOEmpStationCOUNTSTATION.AsString);
          end
          else
          begin
            Memo3.Lines.Add(ADOEmpStationSTATION.AsString + '  ' + ADOEmpStationCOUNTSTATION.AsString);
          end;
          next;
        end;
      end;
    end;

 with Series11 do
 begin
    i := i + 66800;
    if (length(ADOGetNamelastnam.AsString) < 2) then
    next
    else
    begin
      emp := copy(ADOGetNamefstnam.AsString,1,1);
      emp := concat(emp,' ');
      empl := copy(ADOGetNamelastnam.AsString,1,4);
      emp := concat(emp,empl);
      ordcnt := ADOAllEmpsCOUNTEMPLOYEE.AsInteger;
       Add( ordcnt, emp, clRed + i );
       next;
    end;
 end;
  end;
  Memo3.Visible := true;
  screen.Cursor := crDefault;
end;
end;


procedure TForm3.BtnExpOrdClick(Sender: TObject);
    var
  i: integer;
  strT: string;
  slst: TStringList;
begin
  slst:= TStringList.Create;
Try
with dsemp.DataSet do
  begin
        First;
    while not Eof do
    begin
      strT:= '"'+Fields[0].AsString+'"';
      for i:= 1 to FieldCount-1 do
        strT:= strT+',"'+Fields[i].AsString+'"';
      slst.Add(strT);
      Next;
    end;
        First;
  end;
  slst.SaveToFile('C:\Data\Employee.csv');
Finally
  slst.Free;
end;
end;


procedure TForm3.BtnExportClick(Sender: TObject);
 var
i: integer;
strT: string;
slst: TStringList;
var beg, type1:String;
var bDts, eDts: String;
begin
slst:= TStringList.Create;

bDts := FormatDateTime('dd-MMM-yy',DateTimePickerDupBeg.DateTime);
bDts := UpperCase(bDts);

eDts := FormatDateTime('dd-MMM-yy',DateTimePickerDupEnd.DateTime);
eDts := UpperCase(eDts);

with ADODupChk do
     begin
        Close;
        ADODupChk.Parameters.ParamByName('dtsbeg').Value := bDts;
        ADODupChk.Parameters.ParamByName('dtsend').Value := eDts;
        Open;
        First;
        Memo1.Clear;
        Memo1.Font.Size := 12;
        Memo1.Lines.Add('ORDER   EMPLOYEE   STATION   DATE');
        Memo1.Lines.Add('');
        Memo1.Lines.Add('');
       while not eof do
       begin
          with ADODspDup do
          begin
            Close;
            Parameters.ParamByName('orderno').Value := ADODupChkORDERNO.AsString;
            Open;
            while not eof do
            begin

             Memo1.Font.Size := 10;
              Memo1.Lines.Add(ADODspDupORDERNO.AsString + '       '
              + ADODspDupEMPLOYEE.AsString + '      '
              + ADODspDupSTATION.AsString+ '      '
              + ADODspDupDTS.AsString);

           strT:= '"'+Fields[0].AsString+'"';
            for i:= 1 to FieldCount-1 do
            strT:= strT+',"'+Fields[i].AsString+'"';
            slst.Add(strT);
            next;
            end;
          end;
           ADODupChk.next;
       end;
     end;
  slst.SaveToFile('C:\Data\Duplicate.csv');
  slst.Free;
end;


procedure TForm3.Button4Click(Sender: TObject);
var
  i: integer;
  strT: string;
  slst: TStringList;
begin
  slst:= TStringList.Create;
Try
with dsOrder.DataSet do
  begin
        First;
    while not Eof do
    begin
      strT:= '"'+Fields[0].AsString+'"';
      for i:= 1 to FieldCount-1 do
        strT:= strT+',"'+Fields[i].AsString+'"';
      slst.Add(strT);
      Next;
    end;
        First;
  end;
  slst.SaveToFile('C:\Data\Order.csv');
Finally
  slst.Free;
End;
end;

procedure TForm3.Button5Click(Sender: TObject);
var dtsB,dtsE, type1:String;
var amount: Integer;
begin
  Series1.Clear;

  dtsB := FormatDateTime('dd-MMM-yy',DateTimePickerTypeB.DateTime);
  dtsB := UpperCase(dtsB);

  dtsE := FormatDateTime('dd-MMM-yy',DateTimePickerTypeE.DateTime);
  dtsE := UpperCase(dtsE);

with ADOTriage do
begin
  ADOTriage.Close;
  ADOTriage.Parameters.ParamByName('dtsb').Value := dtsB;
  ADOTriage.Parameters.ParamByName('dtse').Value := dtsE;
  ADOTriage.Open;
  First;
  while not eof do
  begin
    with Series1 do
    begin
      amount := ADOTriageCOUNTORDERTYPE.AsInteger;
      type1 := ADOTriageORDERTYPE.AsString;
      if type1 = 'I.V. ' then
      Add( amount, type1 , clRed )
      else if type1 = 'Entral ' then
      Add( amount, type1 , clBlue )
      else if type1 = 'Incontinence' then
      Add( amount, type1 , clGreen )
      else if type1 = 'Will Call' then
      Add( amount, type1 , clPurple )
      else if type1 = 'Pick-Up' then
      Add( amount, type1 , clYellow )
      else if type1 = 'resupply' then
      Add( amount, type1 , clLime )
//      else
//      Add(  amount, 'Undefined' , clNavy ) ;
    end;
     next;
  end;
end;
end;



procedure TForm3.Button8Click(Sender: TObject);
var dts, dtsB,dtsE,sortdts: String;
begin
  dtsB := FormatDateTime('dd-MMM-yy',DateTimePickerBeg.DateTime);
  dtsB := UpperCase(dtsB);

  dtsE := FormatDateTime('dd-MMM-yy',DateTimePickerEnd.DateTime);
  dtsE := UpperCase(dtsE);

  sortdts := sortbydts;

  screen.Cursor := crHourGlass;
  with ADOOrdByDt do
        begin
          Close;
          sql.Clear;
          sql.Add('Select * from mcca.hc_whs');
          sql.Add('where ndts > ''' + dtsB + '''');
          sql.Add('and ndts < ''' + dtsE + '''');
          if  sortdts = '0' then
          sql.Add('order by orderno')
          else if sortdts = '1' then
          sql.Add('order by employee')
          else if sortdts = '2' then
          sql.Add('order by station')
          else if sortdts = '3' then
          sql.Add('order by ordertype')
          else if sortdts = '4' then
          sql.Add('order by dlvry_type')
          else if sortdts = '5' then
          sql.Add('order by pkg_cnt')
          else if sortdts = '6' then
          sql.Add('order by ndts')
          else
          sql.Add('order by orderno');
          Open;
        end;

  with ADODisOrdByDts do
      begin
          ADODisOrdByDts.Close;
          ADODisOrdByDts.Parameters.ParamByName('date3').Value := dtsB;
          ADODisOrdByDts.Parameters.ParamByName('date4').Value := dtsE;
          ADODisOrdByDts.Open;
        end;
   EdtTotal.Text := IntToStr(ADODisOrdByDts.RecordCount);
   screen.Cursor := crDefault;
end;


procedure TForm3.Button9Click(Sender: TObject);
var
i: integer;
strT: string;
slst: TStringList;
begin
  slst:= TStringList.Create;
try
with dsOrdByDts.DataSet do
  begin
        First;
    while not Eof do
    begin
      strT:= '"'+Fields[0].AsString+'"';
      for i:= 1 to FieldCount-1 do
        strT:= strT+',"'+Fields[i].AsString+'"';
      slst.Add(strT);
      Next;
    end;
        First;
  end;
  slst.SaveToFile('C:\Data\ByDate.csv');
Finally
  slst.Free;
end;
end;

procedure TForm3.ComboBox1Click(Sender: TObject);
var name,i,len: Integer;
begin
   screen.Cursor := crHourGlass;
   CbName := ComboBox1.Text;
   len := length(CbName);
   i:=Pos(',',CbName);
   lsst := copy(CbName,1,i-1);
   frst := copy(CbName,i+1,len);
   Button2Click(Sender);
   screen.Cursor := crDefault;
end;



procedure TForm3.DateTimePickerAllEmpsBChange(Sender: TObject);
var dtsTest: String;
begin
  dtsTest := DateToStr(DateTimePickerAllEmpsB.Date);
    if DateTimePickerAllEmpsB.Date < StrToDate('12/10/2016') then
    begin
      Showmessage('Invalid before Dec.10, 2016');
      exit;
    end;
end;

procedure TForm3.DateTimePickerBegChange(Sender: TObject);
var dtsTest: String;
begin
    dtsTest := DateToStr(DateTimePickerBeg.Date);
    if DateTimePickerBeg.Date < StrToDate('12/10/2016') then
    begin
      Showmessage('Invalid before Dec.10, 2016');
      exit;
    end;

end;

procedure TForm3.DateTimePickerEmpBegChange(Sender: TObject);
var dtsTest: String;
begin
   dtsTest := DateToStr(DateTimePickerEmpBeg.Date);
    if DateTimePickerEmpBeg.Date < StrToDate('12/10/2016') then
    begin
      Showmessage('Invalid before Dec.10, 2016');
      exit;
    end;
end;

procedure TForm3.DateTimePickerQCbChange(Sender: TObject);
var dtsTest: String;
begin
  dtsTest := DateToStr(DateTimePickerQCb.Date);
    if DateTimePickerQCb.Date < StrToDate('12/10/2016') then
    begin
      Showmessage('Invalid before Dec.10, 2016');
      exit;
    end;
end;

procedure TForm3.DateTimePickerTQCbChange(Sender: TObject);
var dtsTest: String;
begin
  dtsTest := DateToStr(DateTimePickerTQCb.Date);
    if DateTimePickerTQCb.Date < StrToDate('12/10/2016') then
    begin
      Showmessage('Invalid before Dec.10, 2016');
      exit;
    end;
end;

procedure TForm3.DateTimePickerTypeBChange(Sender: TObject);
var dtsTest: String;
begin
  dtsTest := DateToStr(DateTimePickerTypeB.Date);
    if DateTimePickerTypeB.Date < StrToDate('12/10/2016') then
    begin
      Showmessage('Invalid before Dec.10, 2016');
      exit;
    end;
end;

procedure TForm3.DateTimeStatBChange(Sender: TObject);
var dtsTest: String;
begin
  dtsTest := DateToStr(DateTimeStatB.Date);
    if DateTimeStatB.Date < StrToDate('12/10/2016') then
    begin
      Showmessage('Invalid before Dec.10, 2016');
      exit;
    end;
end;

procedure TForm3.DBGrid2TitleClick(Column: TColumn);
var Sender: TObject;
begin
   case Column.Index of
  0: sort := '0';
  1: sort := '1';
  2: sort := '2';
  3: sort := '3';
  4: sort := '4';
  5: sort := '5';
   end;
   screen.Cursor := crHourGlass;
   Button2Click(Sender);
   screen.Cursor := crDefault;
end;

procedure TForm3.DBGrid3CellClick(Column: TColumn);
var empno: String;
begin
    empno := Column.Field.Value;
    if (length(empno) = 5) or (length(empno) = 6) then
    begin
    with ADOName do
     begin
       Close;
       ADOName.Parameters.ParamByName('empno').Value := empno;
       Open;
     end;
     Showmessage(format('%s', [ADONamelastnam.AsString + ', '+ ADONamefstnam.AsString]));
    end;
end;

procedure TForm3.DBGrid3TitleClick(Column: TColumn);
var Sender: TObject;
begin
   case Column.Index of
  0: sortbydts := '0';
  1: sortbydts := '1';
  2: sortbydts := '2';
  3: sortbydts := '3';
  4: sortbydts := '4';
  5: sortbydts := '5';
  6: sortbydts := '6';
   end;
   screen.Cursor := crHourGlass;
   Button8Click(Sender);
   screen.Cursor := crDefault;

end;


procedure TForm3.DBGrid5CellClick(Column: TColumn);
var empno, type1: String;
var typeindex, position : Integer;
begin
    empno := Column.Field.Value;
    position := ansipos('/', empno);
    if not(position=0) then exit;
    if Column.Field.Value = '' then exit;
    if (empno[1] in ['a'..'z', 'A'..'Z']) and (length(empno) < 6)then
    begin
      typeindex := Column.Index - 1;
      type1 := DBGrid5.Fields[typeindex].Text;
      with GetDesc do
      begin
       Close;
       GetDesc.Parameters.ParamByName('type').Value := type1;
       GetDesc.Parameters.ParamByName('code').Value := empno;
       Open;
      end;
     Showmessage(format('%s', [GetDescERROR_DSC.AsString]))
    end
    else if (length(empno) = 5) or (length(empno) = 6) then
      begin
        with ADOName do
        begin
          Close;
          ADOName.Parameters.ParamByName('empno').Value := empno;
          Open;
        end;
        Showmessage(format('%s', [ADONamelastnam.AsString + ', '+ ADONamefstnam.AsString]));
      end
      else
      exit;
  end;


procedure TForm3.DBGrid5TitleClick(Column: TColumn);
var Sender: TObject;
begin
  case Column.Index of
  0: sortbydts := '0';
  1: sortbydts := '1';
  2: sortbydts := '2';
  3: sortbydts := '3';
  4: sortbydts := '4';
  5: sortbydts := '5';
   end;
   screen.Cursor := crHourGlass;
   BtnErrorClick(Sender);
   screen.Cursor := crDefault;
end;


procedure TForm3.FormCreate(Sender: TObject);
var test:String;
var i: Integer;
begin


//  TStringGrid(DBGrid1).RowHeights[0] := 25;
    PageControl1.ActivePageIndex := 1;
    sort := '0';
    i := 0;
    sortbydts := '0';
    DateTimePickerTypeB.DateTime := date;
    DateTimePickerTypeE.DateTime := date + 1;
    DateTimePickerBeg.DateTime := date;
    DateTimePickerEnd.DateTime := date + 1;
    DateTimePickerQCb.DateTime := date;
    DateTimePickerQCe.DateTime := date + 1;
    DateTimeStatB.DateTime := date;
    DateTimeStatE.DateTime := date + 1;
    DateTimePickerDupBeg.DateTime := date;
    DateTimePickerDupEnd.DateTime := date + 1;
    DateTimePickerAllEmpsB.DateTime := date;
    DateTimePickerAllEmpsE.DateTime := date + 1;
    DateTimePickerTQCb.DateTime := date;
    DateTimePickerTQCe.DateTime := date + 1;
    DateTimePickerEmpBeg.DateTime := date;
    DateTimePickerEmpEnd.DateTime := date + 1;
    DateTimeErrBeg.DateTime := date;
    DateTimeErrEnd.DateTime := date + 1;
    with ADOAmyTm do
    begin
      Close;
      Open;
    while not eof do
    begin
    ComboBox1.Items.Add(ADOAmyTmLASTNAM.AsString + ',' + ADOAmyTmFSTNAM.AsString);
    next;
    end;
    end;
end;




procedure TForm3.Memo4KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  begin
    Memo4.SelStart := 10;
    Memo4.Perform(EM_ScrollCaret, 3, 0);
    Memo1.SetFocus;
  end;

end;

procedure TForm3.PageControl1Change(Sender: TObject);
begin
  if PageControl1.ActivePage.Caption = 'Triage' then
    Button5Click(Sender);
  if PageControl1.ActivePage.Caption = 'Packing/QC' then
    BtnQCClick(Sender);
  if PageControl1.ActivePage.Caption = 'Station' then
    BtnStationClick(Sender);
  if PageControl1.ActivePage.Caption = 'Duplicate Orders' then
    BtnDupClick(Sender);
  if PageControl1.ActivePage.Caption = 'Orders By Date' then
     Button8Click(Sender);
  if PageControl1.ActivePage.Caption = 'Employee Orders' then
    begin
      Memo3.Visible := false;
      BtnEmpOrdsClick(Sender);
    end;
//  if PageControl1.ActivePage.Caption = 'Order Sequence' then
//      BtnInComplClick(Sender);

end;

end.
