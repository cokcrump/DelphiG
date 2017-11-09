unit Labels;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, ADODB, ExtCtrls, Buttons, Keyboard, Grids, DBGrids,
  DBTables, ComCtrls, Printers, WinSpool, shellapi, CheckLst;


type
  TForm1 = class(TForm)
    QryName: TADOQuery;
    QryNameLast_Name: TStringField;
    QryNamedelivdate: TDateTimeField;
    QryNamefirst: TStringField;
    ADOInsWhs: TADOQuery;
    LblProcess: TLabel;
    ADODupChk: TADOQuery;
    DataSource1: TDataSource;
    ADODupChkCNT: TFloatField;
    ADODspDup: TADOQuery;
    ADODupChkORDERNO: TFMTBCDField;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Button1: TButton;
    LblContainers: TLabel;
    LblOrderNo: TLabel;
    LblEmp: TLabel;
    LblOrderType: TLabel;
    EdtContainers: TEdit;
    EdtOrder: TEdit;
    PnlOrdMethod: TPanel;
    LblType: TLabel;
    EdtMethod: TEdit;
    EdtEmpoyee: TEdit;
    EdtOrderType: TEdit;
    CbLock: TCheckBox;
    StationRadioGroup: TRadioGroup;
    TabSheet3: TTabSheet;
    Label2: TLabel;
    Label3: TLabel;
    Button2: TButton;
    EdtDescription: TEdit;
    Button3: TButton;
    DBGridDup: TDBGrid;
    EdtDisHstry: TEdit;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    BitBtnOk: TBitBtn;
    BtnSeach: TButton;
    LblEntOrder: TLabel;
    ADOQrder: TADOQuery;
    ADOName: TADOQuery;
    ADONamefstnam: TWideMemoField;
    LblProcessed: TLabel;
    ADOTracking: TADOQuery;
    ADOTrackingshipmentid: TStringField;
    ADOTrackingtouchdate: TDateTimeField;
    LblDisabled: TLabel;
    EdtWorShp: TEdit;
    Label1: TLabel;
    TabSheet4: TTabSheet;
    EdtLastName: TEdit;
    DateStart: TDateTimePicker;
    BtnWorkOrder: TButton;
    ADOGetOrder: TADOQuery;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    DBGrid1: TDBGrid;
    dsOrder: TDataSource;
    EdtCnt: TEdit;
    Label8: TLabel;
    BtnClear: TButton;
    EdtQuanity: TEdit;
    Label9: TLabel;
    ADOGetOrdertickno: TAutoIncField;
    TabSheet5: TTabSheet;
    RGroupPrinters: TRadioGroup;
    TabSheet6: TTabSheet;
    MemoOther: TMemo;
    Label10: TLabel;
    EdtErrOrder: TEdit;
    Label11: TLabel;
    Label12: TLabel;
    BtnErrors: TButton;
    EdtErrEmp: TEdit;
    ADOInsErrCodes: TADOQuery;
    BtnClearErrors: TButton;
    ADOChkErrs: TADOQuery;
    ADOQrderORDERNO: TFMTBCDField;
    ADOQrderSTATION: TStringField;
    ADOQrderEMPLOYEE: TStringField;
    ADOQrderORDERTYPE: TStringField;
    ADOQrderDTS: TStringField;
    ADOQrderPKG_CNT: TStringField;
    ADOQrderDLVRY_TYPE: TStringField;
    ADOQrderNDTS: TDateTimeField;
    ADOQrderShipId: TStringField;
    ADOQrderShipDts: TStringField;
    DbGridErrors: TDBGrid;
    DataSource2: TDataSource;
    ADONamelastnam: TWideMemoField;
    ADOQrderName: TStringField;
    ADOChkErrsEMPLOYEE: TStringField;
    ADOChkErrsERROR_TYPE: TStringField;
    ADOChkErrsERROR_DSC: TStringField;
    ADOChkErrsNDTS: TDateTimeField;
    ADOChkErrsERROR_DSC_1: TStringField;
    ADOChkErrsName: TStringField;
    ListBoxTriage: TCheckListBox;
    ListBoxRePrints: TCheckListBox;
    ListBoxPutAway: TCheckListBox;
    ListBoxPharmacy: TCheckListBox;
    ListBoxQC: TCheckListBox;
    ListBoxRetail: TCheckListBox;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    ListBoxPicking: TCheckListBox;
    ListBoxNoErrors: TCheckListBox;
    ADOQrderDescription: TStringField;
    RadioButton1: TRadioButton;
    procedure Button1Click(Sender: TObject);
    procedure PrintUSB(OrderId,Last,First,Dts: String; cnt,Cntrns: Integer);
    procedure PrintPrlPrt(OrderId,Last,First,Dts: String; cnt,Cntrns: Integer);

    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure EdtOrderKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure EdtContainersKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure InsertInfo(OrderId: String; Station: String; employee: String; ordertype: String; OrderMethod: String; Cntrns: String);
    procedure EdtMethodKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure StationRadioGroupClick(Sender: TObject);
    procedure EdtOrderTypeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure EdtEmpoyeeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Button3Click(Sender: TObject);
    procedure CbLockClick(Sender: TObject);
    procedure BitBtnOkClick(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure BtnSeachClick(Sender: TObject);
    procedure ADOQrderCalcFields(DataSet: TDataSet);
    procedure BtnWorkOrderClick(Sender: TObject);
    procedure BtnClearClick(Sender: TObject);
    procedure RGroupPrintersClick(Sender: TObject);
    procedure EdtErrOrderKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BtnErrorsClick(Sender: TObject);
    procedure MemoOtherChange(Sender: TObject);
    procedure MemoOtherKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BtnClearErrorsClick(Sender: TObject);
    procedure ADOChkErrsCalcFields(DataSet: TDataSet);
    procedure RadioButton1Click(Sender: TObject);
    procedure ListBoxTriageDrawItem(Control: TWinControl; Index: Integer;
      Rect: TRect; State: TOwnerDrawState);
    procedure EdtErrEmpKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);



  private
    Last : String;
    First : String;
    OrderType : String;
    OrderMethod : String;
    KeepType,USB: Boolean;
    Data : String;
    Code: String;
    etype: String;
    IgnoreWarning: Boolean;
    ErrorsSaved: Boolean;
    overide: Boolean;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

//{$DEFINE GX420}
//{$DEFINE ZP450}
//{$DEFINE PPORT}

procedure TForm1.ADOChkErrsCalcFields(DataSet: TDataSet);
var fullname: string;
begin
     with ADOName do
     begin
       Close;
       ADOName.Parameters.ParamByName('empno').Value := DataSet['employee'];
       Open;
     end;
     fullname :=  Concat(Copy(ADONamefstnam.AsString,1,1),'. ');
     fullname :=  Concat(fullname,ADONamelastnam.AsString);
     DataSet['Name'] := fullname;
end;

procedure TForm1.ADOQrderCalcFields(DataSet: TDataSet);
var fullname: string;
begin
     with ADOName do
     begin
       Close;
       ADOName.Parameters.ParamByName('empno').Value := DataSet['employee'];
       Open;
     end;
     fullname :=  Concat(Copy(ADONamefstnam.AsString,1,1),'. ');
     fullname :=  Concat(fullname,ADONamelastnam.AsString);
     DataSet['Name'] := fullname;

     with ADOTracking do
     begin
       Close;
       ADOTracking.Parameters.ParamByName('orderno').Value := DataSet['ORDERNO'];
       Open;
     end;
     DataSet['ShipId'] := ADOTrackingshipmentid.AsString;
     DataSet['ShipDts'] := ADOTrackingtouchdate.AsString;
end;

procedure TForm1.BitBtnOkClick(Sender: TObject);
begin
    PageControl1.ActivePage.PageControl.ActivePageIndex:= 0;
    LblProcessed.Caption := '';
    if CbLock.Checked then
    EdtOrderType.SetFocus
    else
    EdtEmpoyee.SetFocus;
end;

procedure TForm1.BtnClearClick(Sender: TObject);
begin
  EdtLastName.Text := '';
  DateStart.DateTime := date;
  EdtCnt.Text := '';
  EdtQuanity.Text := '';
  dsOrder.DataSet.Active := False;
  EdtLastName.SetFocus;
end;

procedure TForm1.BtnClearErrorsClick(Sender: TObject);
var i,j: Integer;
begin
   ListBoxTriage.CheckAll(cbUnchecked, true, false);
   ListBoxRePrints.CheckAll(cbUnchecked, true, false);
   ListBoxPicking.CheckAll(cbUnchecked, true, false);
   ListBoxPutAway.CheckAll(cbUnchecked, true, false);
   ListBoxPharmacy.CheckAll(cbUnchecked, true, false);
   ListBoxQC.CheckAll(cbUnchecked, true, false);
   ListBoxRetail.CheckAll(cbUnchecked, true, false);
   ListBoxNoErrors.CheckAll(cbUnchecked, true, false);
   MemoOther.Text := '';
   EdtErrOrder.Text := '';
   EdtErrEmp.Text := '';

    for j:= 0 to ListBoxTriage.Items.Count - 1 do
                if ListBoxTriage.Selected[j] then
                   ListBoxTriage.Selected[j] := False;
   for j:= 0 to ListBoxRePrints.Items.Count - 1 do
                if ListBoxRePrints.Selected[j] then
                   ListBoxRePrints.Selected[j] := False;
   for j:= 0 to ListBoxPicking.Items.Count - 1 do
                if ListBoxPicking.Selected[j] then
                   ListBoxPicking.Selected[j] := False;
   for j:= 0 to ListBoxPutAway.Items.Count - 1 do
                if ListBoxPutAway.Selected[j] then
                   ListBoxPutAway.Selected[j] := False;
   for j:= 0 to ListBoxPharmacy.Items.Count - 1 do
                if ListBoxPharmacy.Selected[j] then
                   ListBoxPharmacy.Selected[j] := False;
   for j:= 0 to ListBoxQC.Items.Count - 1 do
                if ListBoxQC.Selected[j] then
                   ListBoxQC.Selected[j] := False;
  for j:= 0 to ListBoxRetail.Items.Count - 1 do
                if ListBoxRetail.Selected[j] then
                   ListBoxRetail.Selected[j] := False;
  for j:= 0 to ListBoxNoErrors.Items.Count - 1 do
                if ListBoxNoErrors.Selected[j] then
                   ListBoxNoErrors.Selected[j] := False;
end;

procedure TForm1.BtnSeachClick(Sender: TObject);
begin
  screen.Cursor := crHourGlass;
  DbGridErrors.Visible := false;
  with ADOQrder do
    begin
      ADOQrder.Close;
      ADOQrder.Parameters.ParamByName('orderno').Value := EdtDisHstry.Text;
      ADOQrder.Open;
    end;
    if ADOQrder.RecordCount < 1 then
    begin
      Showmessage('No Records Found');
      screen.Cursor := crDefault;
      EdtDisHstry.SetFocus;
      exit;
    end;
    EdtWorShp.Text :=  ADOQrderShipDts.Value;

    with ADOChkErrs do
    begin
      ADOChkErrs.Close;
      ADOChkErrs.Parameters.ParamByName('ordernum').Value := EdtDisHstry.Text;
      ADOChkErrs.Open;
    end;
    if ADOChkErrs.RecordCount < 1 then
      DbGridErrors.Visible := false
    else
     DbGridErrors.Visible := true;
    screen.Cursor := crDefault;
end;

procedure TForm1.BtnWorkOrderClick(Sender: TObject);
var dtsE: String;
var count: Integer;
begin

  dtsE := FormatDateTime('yyyy-mm-dd',DateStart.DateTime);
  screen.Cursor := crHourGlass;
  with ADOGetOrder do
    begin
      ADOGetOrder.Close;
      ADOGetOrder.Parameters.ParamByName('lname').Value := EdtLastName.Text;
      ADOGetOrder.Parameters.ParamByName('start').Value := dtsE;
      ADOGetOrder.Parameters.ParamByName('qty').Value := EdtQuanity.Text;
      ADOGetOrder.Open;
    end;
    count := ADOGetOrder.RecordCount;
    EdtCnt.Text := IntToStr(count);
    if ADOGetOrdertickno.AsString = '' then
    begin
    Showmessage('No Records Found');
    screen.Cursor := crDefault;
    EdtLastName.SetFocus;
      exit;
    end;
//    EdtOrderNo.Text :=  ADOGetOrdertickno.AsString;
  screen.Cursor := crDefault;

end;

procedure TForm1.Button1Click(Sender: TObject);
var Cntrns : Integer;
var I : Integer;
var cnt : Integer;
var OrderId : String;
var Dts, Img : String;
var Station, employee: String;
var CurrentDts: TDateTime;
begin
     screen.Cursor := crHourGlass;
     LblProcessed.Caption := '';
     Cntrns := 0;
     cnt := 1;
     if EdtMethod.Text = 'Delivery' then OrderMethod := 'D'
    else if  EdtMethod.Text = 'UPS' then OrderMethod := 'U'
    else if  EdtMethod.Text = 'Courier' then OrderMethod := 'C'
    else
    OrderType := 'Delivery';

    Employee := EdtEmpoyee.Text;
    CurrentDts := Date;
    OrderId := EdtOrder.Text;

    if (StationRadioGroup.ItemIndex = -1) then
          begin
              ShowMessage('Station field must be populated.');
              Exit;
          end;
     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Triage' then
        begin
            Station := 'Triage';
            employee := EdtEmpoyee.Text;
            ordertype := EdtOrderType.Text;
            OrderMethod := '';
            Cntrns := 0;
        end;
     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Assign' then
        begin
            Station := 'Assign';
            ordertype := '';
            OrderMethod := '';
            Cntrns := 0;
        end;
     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Staging' then
        begin
            Station := 'Staging';
            ordertype := '';
            OrderMethod := '';
            Cntrns := 0;
        end;

     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Dispatch' then
        begin
            Station := 'Dispatch';
            ordertype := '';
            OrderMethod := '';
            Cntrns := 0;
        end;

     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'QC' then
     begin
        with QryName do
        begin
          QryName.Close;
          QryName.Parameters.ParamByName('Number').Value := EdtOrder.Text;
          QryName.Open;
        end;
        Station := 'QC';
        ordertype := '';
        Last := QryNameLast_Name.asString;
        First := QryNamefirst.asString;
        Cntrns := StrToInt(EdtContainers.text);
        Dts := QryNamedelivdate.AsString;
        for I := 1 to Cntrns do
          begin
            sleep(500);
             if RGroupPrinters.ItemIndex = 0 then
             PrintUSB(OrderId,Last,First,Dts,cnt,Cntrns)
             else if  RGroupPrinters.ItemIndex = 1 then
             PrintPrlPrt(OrderId,Last,First,Dts,cnt,Cntrns)
             else  PrintUSB(OrderId,Last,First,Dts,cnt,Cntrns);
            Inc(cnt);
        end;
     end;


     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Triage' then
        begin
            if CbLock.Checked then
              EdtOrderType.SetFocus
            else
              EdtEmpoyee.SetFocus;
        end;
     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Assign' then
            EdtEmpoyee.SetFocus;
     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Staging' then
            EdtEmpoyee.SetFocus;
     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Dispatch' then
            EdtEmpoyee.SetFocus;
     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'QC' then
        begin
            if CbLock.Checked then
              EdtMethod.SetFocus
            else
              EdtEmpoyee.SetFocus;
        end;
     InsertInfo(OrderId,Station,employee,ordertype,OrderMethod,IntToStr(Cntrns));
     LblProcessed.Caption := 'Processed';
     screen.Cursor := crDefault;
end;



procedure TForm1.MemoOtherChange(Sender: TObject);
begin
    if (MemoOther.Lines.Count > 2) and (IgnoreWarning = False) then
    begin
      if MessageDlg ('Two row maxium, please finish and exit' + #10#13 +
          'Continue displaying this warning?',
          mtWarning, [mbYes, mbNo], 0) = mrNo then
          IgnoreWarning := True;
    end;
end;

procedure TForm1.MemoOtherKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if (MemoOther.CaretPos.X > 70) and not ((key=VK_RETURN) or (key=VK_BACK)
    or (key=VK_UP) or (key=VK_DOWN) or (key=VK_LEFT) or (key=VK_RIGHT) or (key = VK_HOME))
    then
    begin
    ShowMessage('Maxium line length Exceeded !!!')
    end;
end;


procedure TForm1.InsertInfo(orderId: String; station: String; employee: String; ordertype: String; ordermethod: String; cntrns: String);
var sFileName, sCurDir: String;
var FSize : longint;
var  f : textfile;
var extension, file_name: String;
var server_name, Dts: String;
var CurrentTime: String;
var CurrentDate: String;
begin


with ADOInsWhs do
begin
      Close;
      Parameters.ParamByName('ORDERID').Value := OrderId;
      Parameters.ParamByName('STATION').Value := Station;
      Parameters.ParamByName('EMPLOYEE').Value := employee;
      Parameters.ParamByName('STATION').Value := Station;
      Parameters.ParamByName('ORDERTYPE').Value := ordertype;
      Parameters.ParamByName('DLVRY_TYPE').Value := ordermethod;
      Parameters.ParamByName('PKG_CNT').Value := cntrns;
      ExecSQL;
end;
end;


procedure TForm1.ListBoxTriageDrawItem(Control: TWinControl; Index: Integer;
  Rect: TRect; State: TOwnerDrawState);
begin
 if odSelected in State then
  begin
  ListBoxTriage.Canvas.Font.Color := clBlack;
  ListBoxTriage.Canvas.Brush.Color := clWhite;
  ListBoxTriage.Canvas.FillRect(Rect);
  ListBoxTriage.Canvas.TextOut(Rect.Left,Rect.Top,ListBoxTriage.Items[index]);
  end;
end;

procedure TForm1.PrintUSB(OrderId,Last,First,Dts: String; cnt,Cntrns: Integer);
var
   Handle: THandle;
   N: DWORD;
   DocInfo1: TDocInfo1;
   sPrinterName : string;
   sTexte : AnsiString;
   sString : AnsiString;
   F: TextFile;
   S: string;
   Index: Integer;
   PrntName: String;
   ALPHA_FO_X: String;
   ALPHA_FO_Y: String;
   ALPHA_AON_X: String;
   ALPHA_AON_Y: String;
   BAR_FO_X: String;
   BAR_FO_Y: String;
   BAR_BCN: String;
   FAC_FO_X :String;
   FAC_FO_Y :String;
   FAC_AON_X :String;
   FAC_AON_Y :String;
   FAC1_FO_X :String;
   FAC1_FO_Y :String;
   FAC1_AON_X :String;
   FAC1_AON_Y :String;
   FAC2_FO_X :String;
   FAC2_FO_Y :String;
   FAC2_AON_X :String;
   FAC2_AON_Y :String;
   FAC3_FO_X :String;
   FAC3_FO_Y :String;
   FAC3_AON_X :String;
   FAC3_AON_Y :String;
   FAC4_FO_X :String;
   FAC4_FO_Y :String;
   FAC4_AON_X :String;
   FAC4_AON_Y :String;
   FAC5_FO_X :String;
   FAC5_FO_Y :String;
   FAC5_AON_X :String;
   FAC5_AON_Y :String;
   FAC6_FO_X :String;
   FAC6_FO_Y :String;
   FAC6_AON_X :String;
   FAC6_AON_Y :String;
   ALPHA_FO_X1 :String;
   ALPHA_FO_Y1 :String;
   ALPHA_AON_X1 :String;
   ALPHA_AON_Y1: String;
   Delivery_X :String;
   Delivery_Y :String;
   Delivery_AON_X :String;
   Delivery_AON_Y :String;

begin

      AssignFile(F,'Document1.txt');
      ReWrite(F);


// First Name
//          ALPHA_FO_X := ('140');  // First Name
          ALPHA_FO_X := ('50');  // First Name
          ALPHA_FO_Y :=  ('34'); // First Name
          ALPHA_AON_X := ('75');  // First Name
          ALPHA_AON_Y := ('75');  // First Name length

// Peroid
//          ALPHA_FO_X1 := ('200');  // First Name
          ALPHA_FO_X1 := ('110');  // Period Name
          ALPHA_FO_Y1 :=  ('34'); // First Name
          ALPHA_AON_X1 := ('75');  // First Name
          ALPHA_AON_Y1:= ('75');  // First Name length

// Last Name

//          BAR_FO_X :=  ('220');  // Last Name
           BAR_FO_X :=  ('130');  // Last Name
          BAR_FO_Y := ('34');  // Last Name y position
          BAR_BCN := ('75');   // Last Name size

// Order Number

//          FAC_FO_X := ('130');   // Order position
          FAC_FO_X := ('40');   // Order position
          FAC_FO_Y :=  ('334');  // Order
          FAC_AON_X := ('75');   // Order
          FAC_AON_Y := ('75');  // Order size

// Date
//          FAC4_FO_X := ('400');
          FAC4_FO_X := ('310');
          FAC4_FO_Y :=  ('334');
          FAC4_AON_X := ('75');
          FAC4_AON_Y := ('65');

// Total Containers

//          FAC1_FO_X := ('370');
          FAC1_FO_X := ('280');
          FAC1_FO_Y :=  ('160');
          FAC1_AON_X := ('150');
          FAC1_AON_Y := ('82');

// Count Containers

//          FAC2_FO_X := ('140');   // Position x
          FAC2_FO_X := ('50');   // Position x
          FAC2_FO_Y :=  ('160');
          FAC2_AON_X := ('150');
          FAC2_AON_Y := ('82');

// Of
//          FAC3_FO_X := ('250');
          FAC3_FO_X := ('160');
          FAC3_FO_Y :=  ('160');
          FAC3_AON_X := ('150');
          FAC3_AON_Y := ('82');


// Delivery

//          Delivery_X := ('550');
          Delivery_X := ('460');
          Delivery_Y :=  ('230');
          Delivery_AON_X := ('30');
          Delivery_AON_Y := ('30');




       ReWrite(F);
       writeln( F,'^XA^ID*.*^XZ');
       writeln( F,'^XA^PRC');
       writeln( F,'^LH0,0^FS');
       writeln( F,'^LL591');
       writeln( F,'^MD0');
       writeln( F,'^LH0,0^FS');

       writeln( F,'^FO' + ALPHA_FO_X + ',' + ALPHA_FO_Y + '^A0N,' + ALPHA_AON_X + ',' + ALPHA_AON_Y + '^CI13^FR^FD' + First + '^FS');         // First Name
       writeln( F,'^FO' + ALPHA_FO_X1 + ',' + ALPHA_FO_Y1 + '^A0N,' + ALPHA_AON_X1 + ',' + ALPHA_AON_Y1 + '^CI13^FR^FD.  ^FS');


       writeln( F,'^FO' + BAR_FO_X + ',' + BAR_FO_y + '^A0N,' + BAR_BCN + '^CI13^FR^FD' + Last + '^FS');                                      // Last Name
       writeln( F,'^FO' + FAC_FO_X + ',' + FAC_FO_Y + '^A0N,' + FAC_AON_X + ',' + FAC_AON_Y + '^CI13^FR^FD' + OrderId + '^FS');     // Order
       writeln( F,'^FO' + FAC4_FO_X + ',' + FAC4_FO_Y + '^A0N,' + FAC4_AON_X + ',' + FAC4_AON_Y + '^CI13^FR^FD' + Dts + '^FS');               // Date

       writeln( F,'^FO' + FAC1_FO_X + ',' + FAC1_FO_Y + '^A0N,' + FAC1_AON_X + ',' + FAC1_AON_Y + '^CI13^FR^FD' + IntToStr(Cntrns) + '^FS');   // Total Contaniers
       writeln( F,'^FO' + FAC2_FO_X + ',' + FAC2_FO_Y + '^A0N,' + FAC2_AON_X + ',' + FAC2_AON_Y + '^CI13^FR^FD' + IntToStr(cnt) + '^FS');      // Container Count
       writeln( F,'^FO' + FAC3_FO_X + ',' + FAC3_FO_Y + '^A0N,' + FAC3_AON_X + ',' + FAC3_AON_Y + '^CI13^FR^FDOF  ^FS');

////////////////////   UPS  ///////////////////////////////////////////////////


if OrderMethod = 'U' then
begin


//writeln( F,'^FO450,115^GFA,6402,6402,33,,::::::::::::::::::::::::::::::::::::gG0gHFE,gG0CP0CP06,');
writeln( F,'^FO360,115^GFA,6402,6402,33,,::::::::::::::::::::::::::::::::::::gG0gHFE,gG0CP0CP06,');
writeln( F,'gG0CP04P06,:gG0CM0FC0407CM06,gG0CL01FE040FEM06,gG0CL0187043C7M06,gG0CL0183C4783M06,');
writeln( F,'gG0CL0181C4F03M06,gG0CL01C0E4E07M06,gG0CM0F075C1EM06,gG0CM07C3F87CM06,gG0CM01F1F1FN06,');
writeln( F,'gG0CN07IFCN06,gG0CN01IFO06,gG0gHFE,:gG0CO0E5EO06,gG0CN03C47O06,gG0CN07843CN06,');
writeln( F,'gG0CN0F041EN06,gG0CM01C0407N06,gG0CM0780403CM06,gG0CM0F00401EM06,gG0CM0C004006M06,');
writeln( F,'gG0CM08004P06,gG0CP04P06,:::gG0gHFE,:,:::::gG0gHFE,:gG0CgG06,::::gG0CQ0CO06,');
writeln( F,'gG0C3F803F83F8FFCI0FFE006,gG0C3F803F83FBFFE001IF806,gG0C3F803F83KF007IFC06,');
writeln( F,'gG0C3F803F83KF80JFE06,gG0C3F803F83FF1FF81JFE06,gG0C3F803F83FE07FC1KF06,');
writeln( F,'gG0C3F803F83FC03FC3FF3FF06,gG0C3F803F83FC03FC3FC0FF86,gG0C3F803F83F803FC3F807F86,');
writeln( F,'gG0C3F803F83F801FC3F807F86,:gG0C3F803F83F801FC3F803F86,gG0C3F803F83F801FC3F8J06,');
writeln( F,'gG0C3F803F83F801FC3FCJ06,gG0C3F803F83F801FC3FEJ06,gG0C3F803F83F801FC3IFI06,');
writeln( F,'gG0C3F803F83F801FC1JF006,gG0C3F803F83F801FC1JF806,gG0C3F803F83F801FC0JFE06,');
writeln( F,'gG0C3F803F83F801FC07JF06,gG0C3F803F83F801FC01JF06,gG0C3F803F83F801FC007IF86,');
writeln( F,'gG0C3F803F83F801FCI07FF86,gG0C3F803F83F801FCJ0FFC6,gG0C3F803F83F801FCJ03FC6,');
writeln( F,'gG0C3F803F83F801FC3F803FC6,:gG0C3F803F83F801FC3F801FC6,gG0C3F803F83F801FC3F803FC6,');
writeln( F,'gG0C3FC03F83F803FC3FC03FC6,gG0C3FE07F83FC03FC3FE07FC6,gG081FF1FF83FC07F81FF1FFC6,');
writeln( F,'gG0C1KF03FF0FF81KF86,gG0C0JFE03KF80KF06,gG0C07IFE03KF00KF06,gG0C03IFC03KF007IFE06,');
writeln( F,'gG0C01IF803JFE003IFC04,gG0C00FFE003F8FF8I0IF004,gG04I0FI03F8M0F8004,gG04M03F8Q0C,');
writeln( F,'gG06M03F8Q0C,:gG06M03F8Q08,gG03M03F8P018,:gG018L03F8P03,:gH0CL03F8P06,gH0CL03F8P0E,');
writeln( F,'gH06L03F8P0C,gH03L03F8O018,gH038K03F8O03,gH01CK03F8O06,gI0EW0C,gI07V018,gI01CU07,');
writeln( F,'gJ0EU0E,gJ078S038,gJ01CS07,gK0FR01C,gK038Q078,gL0EP01E,gL038O038,gM0EO0E,gM038M0381,');
writeln( F,'gM01EM0E,gN078K038,gN01EJ01EI0A,gO078I038,gP0E001EJ028,gP07C078I02,gP01F1EK08,');
writeln( F,'gQ03F8,gR0E,,::::::::::::::::::::::::::::::::::::^FS');
end;


/////////////////////   Delivery   ////////////////////////////////////////////

if OrderMethod = 'D' then
begin

//writeln( F,'^FO550,150^GFA,1092,1092,12,,::::::::::::::::::::7SFC,::::7SFC7IF8,7SFC7IF,');
writeln( F,'^FO460,150^GFA,1092,1092,12,,::::::::::::::::::::7SFC,::::7SFC7IF8,7SFC7IF,');
writeln( F,':7SFC7IF8,:::7SFC7803C,7SFC7801C,:7SFC7801E,:7SFC7800E,7SFC7800F,7SFC78007,');
writeln( F,':7SFC780078,7SFC780038,7SFC7JFC,7SFC7KF,7SFC7KFE,7SFC7LF8,7SFC7LFC,');
writeln( F,':::::7SFC7FC0IFC,7SF87F003FFC,U07EI0FFC,K01F8M07C3F87FC,7IFC7FC3MF87FC7FC,');
writeln( F,'7IF8FFE3MF8FFE3FC,7IF8F0F1MF1E0E3FC,7IF9E071MF1C071FC,07FF9C079MF1C071FC,');
writeln( F,'J01C038M01807,J01C078M01C07,J01E07N01C07,K0F0FN01E1E,K07FEO0FFE,K07FCO07FC,');
writeln( F,'K01F8O01F,,::::::::::::::::::::^FS');
writeln( F,'^FO' + Delivery_X + ',' + Delivery_Y + '^A0N,' + Delivery_AON_X + ',' + Delivery_AON_Y + '^CI13^FR^FDDELIVERY  ^FS');
end;

/////////////////////   COURIER   /////////////////////////////////////////////////


if OrderMethod = 'C' then
begin

//writeln( F,'^FO550,150^GFA,1152,1152,12,,:::::::::::::::::::::::::::M0MFE,L03NF,L07F801C0038,');
writeln( F,'^FO460,150^GFA,1152,1152,12,,:::::::::::::::::::::::::::M0MFE,L03NF,L07F801C0038,');
writeln( F,'L0FF001C001C,K01FE001CI0C,K01FC001CI0E,K03FC001CI06,K07F8001CI03,K0FF8001CI038,');
writeln( F,'K0FFI01CI018,J01FFI01CI01C,J01FEI01CJ0C,J03FEI01CJ0E,J07MF8I06,J07SF8,001XFC,');
writeln( F,'00gFC,03gFE,03gGF,07gGF,07gGF807FF03RFC0IF80FFC007PFE003FF80FF8FC3PFC3F0FF80FF');
writeln( F,'3FF1PF8FFC7F80FE7FF8PF1FFE7FE0FCIFC7OF3IF3FE7FCF83E7NFE7E1F1FE7F9F01E7NFE780F ');
writeln( F,'9FE7F9E01F3NFE78079FE7F9E00F3NFCF0079FE7F1E00F3NFCF0078FE001E00FP0F0078,');
writeln( F,'001E01FP078078,001F01EP0780F8,I0F83EP07E1F,I0IFCP03IF,I07FF8P01FFE,I03FFR0FFC,');
writeln( F,'J0FCR03F,,:::::::::::::::::::::::::::^FS');
writeln( F,'^FO' + Delivery_X + ',' + Delivery_Y + '^A0N,' + Delivery_AON_X + ',' + Delivery_AON_Y + '^CI13^FR^FDCOURIER  ^FS');
end;

       writeln( F,'^PQ1,0,1,N');
       writeln( F,'^XZ');
       writeln( F,'^FX End of job');
       writeln( F,'^XA');
       writeln( F,'^XZ');


      sPrinterName := Printer.Printers.Strings[Printer.PrinterIndex];

      Caption := sPrinterName;

      if not OpenPrinter(PChar(sPrinterName), Handle, nil) then
      begin
        ShowMessage('error ' + IntToStr(GetLastError));
        Exit;
      end;

      with DocInfo1 do
      begin
        pDocName := PChar('Document1');
        pOutputFile := nil;
        pDataType := 'RAW';
      end;

      Reset(F);
      sString := '';

      while not Eof(F) do
      begin
        ReadLn(F, sTexte);
        sString := sString + sTexte;
      end;

      WinSpool.StartDocPrinter(Handle, 1, @DocInfo1);
      WinSpool.StartPagePrinter(Handle);
      WinSpool.WritePrinter(Handle, PAnsiChar(sString), Length(sString) , N);
      WinSpool.EndPagePrinter(Handle);
      WinSpool.EndDocPrinter(Handle);
      WinSpool.ClosePrinter(Handle);
      CloseFile(F);
end;


procedure TForm1.RadioButton1Click(Sender: TObject);
begin
  PageControl1.ActivePage.PageControl.ActivePageIndex:= 5;
  EdtErrOrder.Text := EdtOrder.Text;
  EdtErrEmp.Text := EdtEmpoyee.Text;
  EdtErrOrder.SetFocus;
  ErrorsSaved := True;
  RadioButton1.Checked := false;
end;

procedure TForm1.RGroupPrintersClick(Sender: TObject);
begin
  if RGroupPrinters.ItemIndex = 0 then
  Caption := 'Version 18     Connection USB'
  else if RGroupPrinters.ItemIndex = 1 then
  Caption := 'Version 18     Connection Parallel Port'
  else
  Caption := 'Version 18     No Connection Type'
end;

procedure TForm1.PageControl1Change(Sender: TObject);
begin
   if (PageControl1.ActivePage.PageControl.ActivePageIndex = 0) and (ErrorsSaved = false) or
   (PageControl1.ActivePage.PageControl.ActivePageIndex = 1) and (ErrorsSaved = false) or
   (PageControl1.ActivePage.PageControl.ActivePageIndex = 2) and (ErrorsSaved = false) or
   (PageControl1.ActivePage.PageControl.ActivePageIndex = 3) and (ErrorsSaved = false) or
   (PageControl1.ActivePage.PageControl.ActivePageIndex = 4) and (ErrorsSaved = false) then
   begin
     PageControl1.ActivePage.PageControl.ActivePageIndex:= 5;
     Showmessage('Error Message Not Saved');
     exit;
   end;



   if PageControl1.ActivePage.Caption = 'Order History' then
   begin
   EdtDisHstry.Visible := true;
   BtnSeach.Visible := true;
   BitBtnOk.Visible := false;
   LblEntOrder.Visible := true;
   DBGridDup.Visible := true;
   LblEntOrder.Caption := 'Order Number';
   DBGridDup.Color := ClNavy;
   EdtDisHstry.SetFocus;
   end;

   if PageControl1.ActivePage.Caption = 'Record Error' then
   begin
      EdtErrOrder.SetFocus;
   end;

   if PageControl1.ActivePage.Caption = 'Input' then
   begin
    if (StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'QC') and (CbLock.Checked = true) then EdtMethod.SetFocus
    else if  (StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'QC') and (CbLock.Checked = false) then EdtEmpoyee.SetFocus
    else if  (StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Triage') and (CbLock.Checked = true) then EdtOrderType.SetFocus
    else if  (StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Triage') and (CbLock.Checked = false) then EdtEmpoyee.SetFocus
    else if  (StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Assign') then EdtEmpoyee.SetFocus
    else if  (StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Staging')then EdtEmpoyee.SetFocus
    else if  (StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Dispatch') then EdtEmpoyee.SetFocus
    else EdtEmpoyee.SetFocus
   end;
end;

procedure TForm1.PrintPrlPrt(OrderId,Last,First,Dts: String; cnt,Cntrns: Integer);
var
   F: TextFile;
   S: string;
   ALPHA_FO_X: String;
   ALPHA_FO_Y: String;
   ALPHA_AON_X: String;
   ALPHA_AON_Y: String;
   BAR_FO_X: String;
   BAR_FO_Y: String;
   BAR_BCN: String;
   FAC_FO_X :String;
   FAC_FO_Y :String;
   FAC_AON_X :String;
   FAC_AON_Y :String;
   FAC1_FO_X :String;
   FAC1_FO_Y :String;
   FAC1_AON_X :String;
   FAC1_AON_Y :String;
   FAC2_FO_X :String;
   FAC2_FO_Y :String;
   FAC2_AON_X :String;
   FAC2_AON_Y :String;
   FAC3_FO_X :String;
   FAC3_FO_Y :String;
   FAC3_AON_X :String;
   FAC3_AON_Y :String;
   FAC4_FO_X :String;
   FAC4_FO_Y :String;
   FAC4_AON_X :String;
   FAC4_AON_Y :String;
   FAC5_FO_X :String;
   FAC5_FO_Y :String;
   FAC5_AON_X :String;
   FAC5_AON_Y :String;
   FAC6_FO_X :String;
   FAC6_FO_Y :String;
   FAC6_AON_X :String;
   FAC6_AON_Y :String;
   ALPHA_FO_X1 :String;
   ALPHA_FO_Y1 :String;
   ALPHA_AON_X1 :String;
   ALPHA_AON_Y1: String;
   Delivery_X :String;
   Delivery_Y :String;
   Delivery_AON_X :String;
   Delivery_AON_Y :String;
begin


// First Name
//          ALPHA_FO_X := ('140');  // First Name
           ALPHA_FO_X := ('50');  // First Name
          ALPHA_FO_Y :=  ('34'); // First Name
          ALPHA_AON_X := ('75');  // First Name
          ALPHA_AON_Y := ('75');  // First Name length


// Peroid
//          ALPHA_FO_X1 := ('200');  // First Name
           ALPHA_FO_X1 := ('110');  // Period Name
          ALPHA_FO_Y1 :=  ('34'); // First Name
          ALPHA_AON_X1 := ('75');  // First Name
          ALPHA_AON_Y1:= ('75');  // First Name length

// Last Name

//          BAR_FO_X :=  ('220');  // Last Name
          BAR_FO_X :=  ('130');  // Last Name
          BAR_FO_Y := ('34');  // Last Name y position
          BAR_BCN := ('75');   // Last Name size

// Order Number

//          FAC_FO_X := ('130');   // Order position
          FAC_FO_X := ('40');   // Order position
          FAC_FO_Y :=  ('334');  // Order
          FAC_AON_X := ('75');   // Order
          FAC_AON_Y := ('75');  // Order size

// Date
//          FAC4_FO_X := ('400');
          FAC4_FO_X := ('310');
          FAC4_FO_Y :=  ('334');
          FAC4_AON_X := ('75');
          FAC4_AON_Y := ('65');

// Total Containers

//          FAC1_FO_X := ('370');
          FAC1_FO_X := ('280');
          FAC1_FO_Y :=  ('160');
          FAC1_AON_X := ('150');
          FAC1_AON_Y := ('82');

// Count Containers

//          FAC2_FO_X := ('140');   // Position x
          FAC2_FO_X := ('50');   // Position x
          FAC2_FO_Y :=  ('160');
          FAC2_AON_X := ('150');
          FAC2_AON_Y := ('82');

// Of
//          FAC3_FO_X := ('250');
          FAC3_FO_X := ('160');
          FAC3_FO_Y :=  ('160');
          FAC3_AON_X := ('150');
          FAC3_AON_Y := ('82');


// Delivery

//          Delivery_X := ('550');
          Delivery_X := ('460');
          Delivery_Y :=  ('230');
          Delivery_AON_X := ('30');
          Delivery_AON_Y := ('30');


       AssignFile(F,'LPT1:');
       ReWrite(F);
       writeln( F,'^XA^ID*.*^XZ');
       writeln( F,'^XA^PRC');
       writeln( F,'^LH0,0^FS');
       writeln( F,'^LL591');
       writeln( F,'^MD0');
       writeln( F,'^LH0,0^FS');

       writeln( F,'^FO' + ALPHA_FO_X + ',' + ALPHA_FO_Y + '^A0N,' + ALPHA_AON_X + ',' + ALPHA_AON_Y + '^CI13^FR^FD' + First + '^FS');         // First Name
       writeln( F,'^FO' + ALPHA_FO_X1 + ',' + ALPHA_FO_Y1 + '^A0N,' + ALPHA_AON_X1 + ',' + ALPHA_AON_Y1 + '^CI13^FR^FD.  ^FS');


       writeln( F,'^FO' + BAR_FO_X + ',' + BAR_FO_y + '^A0N,' + BAR_BCN + '^CI13^FR^FD' + Last + '^FS');                             // Last Name
       writeln( F,'^FO' + FAC_FO_X + ',' + FAC_FO_Y + '^A0N,' + FAC_AON_X + ',' + FAC_AON_Y + '^CI13^FR^FD' + OrderId + '^FS');     // Order
       writeln( F,'^FO' + FAC4_FO_X + ',' + FAC4_FO_Y + '^A0N,' + FAC4_AON_X + ',' + FAC4_AON_Y + '^CI13^FR^FD' + Dts + '^FS');               // Date

       writeln( F,'^FO' + FAC1_FO_X + ',' + FAC1_FO_Y + '^A0N,' + FAC1_AON_X + ',' + FAC1_AON_Y + '^CI13^FR^FD' + IntToStr(Cntrns) + '^FS');   // Total Contaniers
       writeln( F,'^FO' + FAC2_FO_X + ',' + FAC2_FO_Y + '^A0N,' + FAC2_AON_X + ',' + FAC2_AON_Y + '^CI13^FR^FD' + IntToStr(cnt) + '^FS');      // Container Count
       writeln( F,'^FO' + FAC3_FO_X + ',' + FAC3_FO_Y + '^A0N,' + FAC3_AON_X + ',' + FAC3_AON_Y + '^CI13^FR^FDOF  ^FS');

////////////////////   UPS  ///////////////////////////////////////////////////


if OrderMethod = 'U' then
begin


//writeln( F,'^FO450,115^GFA,6402,6402,33,,::::::::::::::::::::::::::::::::::::gG0gHFE,gG0CP0CP06,');
writeln( F,'^FO360,115^GFA,6402,6402,33,,::::::::::::::::::::::::::::::::::::gG0gHFE,gG0CP0CP06,');
writeln( F,'gG0CP04P06,:gG0CM0FC0407CM06,gG0CL01FE040FEM06,gG0CL0187043C7M06,gG0CL0183C4783M06,');
writeln( F,'gG0CL0181C4F03M06,gG0CL01C0E4E07M06,gG0CM0F075C1EM06,gG0CM07C3F87CM06,gG0CM01F1F1FN06,');
writeln( F,'gG0CN07IFCN06,gG0CN01IFO06,gG0gHFE,:gG0CO0E5EO06,gG0CN03C47O06,gG0CN07843CN06,');
writeln( F,'gG0CN0F041EN06,gG0CM01C0407N06,gG0CM0780403CM06,gG0CM0F00401EM06,gG0CM0C004006M06,');
writeln( F,'gG0CM08004P06,gG0CP04P06,:::gG0gHFE,:,:::::gG0gHFE,:gG0CgG06,::::gG0CQ0CO06,');
writeln( F,'gG0C3F803F83F8FFCI0FFE006,gG0C3F803F83FBFFE001IF806,gG0C3F803F83KF007IFC06,');
writeln( F,'gG0C3F803F83KF80JFE06,gG0C3F803F83FF1FF81JFE06,gG0C3F803F83FE07FC1KF06,');
writeln( F,'gG0C3F803F83FC03FC3FF3FF06,gG0C3F803F83FC03FC3FC0FF86,gG0C3F803F83F803FC3F807F86,');
writeln( F,'gG0C3F803F83F801FC3F807F86,:gG0C3F803F83F801FC3F803F86,gG0C3F803F83F801FC3F8J06,');
writeln( F,'gG0C3F803F83F801FC3FCJ06,gG0C3F803F83F801FC3FEJ06,gG0C3F803F83F801FC3IFI06,');
writeln( F,'gG0C3F803F83F801FC1JF006,gG0C3F803F83F801FC1JF806,gG0C3F803F83F801FC0JFE06,');
writeln( F,'gG0C3F803F83F801FC07JF06,gG0C3F803F83F801FC01JF06,gG0C3F803F83F801FC007IF86,');
writeln( F,'gG0C3F803F83F801FCI07FF86,gG0C3F803F83F801FCJ0FFC6,gG0C3F803F83F801FCJ03FC6,');
writeln( F,'gG0C3F803F83F801FC3F803FC6,:gG0C3F803F83F801FC3F801FC6,gG0C3F803F83F801FC3F803FC6,');
writeln( F,'gG0C3FC03F83F803FC3FC03FC6,gG0C3FE07F83FC03FC3FE07FC6,gG081FF1FF83FC07F81FF1FFC6,');
writeln( F,'gG0C1KF03FF0FF81KF86,gG0C0JFE03KF80KF06,gG0C07IFE03KF00KF06,gG0C03IFC03KF007IFE06,');
writeln( F,'gG0C01IF803JFE003IFC04,gG0C00FFE003F8FF8I0IF004,gG04I0FI03F8M0F8004,gG04M03F8Q0C,');
writeln( F,'gG06M03F8Q0C,:gG06M03F8Q08,gG03M03F8P018,:gG018L03F8P03,:gH0CL03F8P06,gH0CL03F8P0E,');
writeln( F,'gH06L03F8P0C,gH03L03F8O018,gH038K03F8O03,gH01CK03F8O06,gI0EW0C,gI07V018,gI01CU07,');
writeln( F,'gJ0EU0E,gJ078S038,gJ01CS07,gK0FR01C,gK038Q078,gL0EP01E,gL038O038,gM0EO0E,gM038M0381,');
writeln( F,'gM01EM0E,gN078K038,gN01EJ01EI0A,gO078I038,gP0E001EJ028,gP07C078I02,gP01F1EK08,');
writeln( F,'gQ03F8,gR0E,,::::::::::::::::::::::::::::::::::::^FS');
end;


/////////////////////   Delivery   ////////////////////////////////////////////

if OrderMethod = 'D' then
begin

//writeln( F,'^FO550,150^GFA,1092,1092,12,,::::::::::::::::::::7SFC,::::7SFC7IF8,7SFC7IF,');
writeln( F,'^FO460,150^GFA,1092,1092,12,,::::::::::::::::::::7SFC,::::7SFC7IF8,7SFC7IF,');
writeln( F,':7SFC7IF8,:::7SFC7803C,7SFC7801C,:7SFC7801E,:7SFC7800E,7SFC7800F,7SFC78007,');
writeln( F,':7SFC780078,7SFC780038,7SFC7JFC,7SFC7KF,7SFC7KFE,7SFC7LF8,7SFC7LFC,');
writeln( F,':::::7SFC7FC0IFC,7SF87F003FFC,U07EI0FFC,K01F8M07C3F87FC,7IFC7FC3MF87FC7FC,');
writeln( F,'7IF8FFE3MF8FFE3FC,7IF8F0F1MF1E0E3FC,7IF9E071MF1C071FC,07FF9C079MF1C071FC,');
writeln( F,'J01C038M01807,J01C078M01C07,J01E07N01C07,K0F0FN01E1E,K07FEO0FFE,K07FCO07FC,');
writeln( F,'K01F8O01F,,::::::::::::::::::::^FS');
writeln( F,'^FO' + Delivery_X + ',' + Delivery_Y + '^A0N,' + Delivery_AON_X + ',' + Delivery_AON_Y + '^CI13^FR^FDDELIVERY  ^FS');
end;

/////////////////////   COURIER   /////////////////////////////////////////////////


if OrderMethod = 'C' then
begin

//writeln( F,'^FO550,150^GFA,1152,1152,12,,:::::::::::::::::::::::::::M0MFE,L03NF,L07F801C0038,');
writeln( F,'^FO460,150^GFA,1152,1152,12,,:::::::::::::::::::::::::::M0MFE,L03NF,L07F801C0038,');
writeln( F,'L0FF001C001C,K01FE001CI0C,K01FC001CI0E,K03FC001CI06,K07F8001CI03,K0FF8001CI038,');
writeln( F,'K0FFI01CI018,J01FFI01CI01C,J01FEI01CJ0C,J03FEI01CJ0E,J07MF8I06,J07SF8,001XFC,');
writeln( F,'00gFC,03gFE,03gGF,07gGF,07gGF807FF03RFC0IF80FFC007PFE003FF80FF8FC3PFC3F0FF80FF');
writeln( F,'3FF1PF8FFC7F80FE7FF8PF1FFE7FE0FCIFC7OF3IF3FE7FCF83E7NFE7E1F1FE7F9F01E7NFE780F ');
writeln( F,'9FE7F9E01F3NFE78079FE7F9E00F3NFCF0079FE7F1E00F3NFCF0078FE001E00FP0F0078,');
writeln( F,'001E01FP078078,001F01EP0780F8,I0F83EP07E1F,I0IFCP03IF,I07FF8P01FFE,I03FFR0FFC,');
writeln( F,'J0FCR03F,,:::::::::::::::::::::::::::^FS');
writeln( F,'^FO' + Delivery_X + ',' + Delivery_Y + '^A0N,' + Delivery_AON_X + ',' + Delivery_AON_Y + '^CI13^FR^FDCOURIER  ^FS');
end;

       writeln( F,'^PQ1,0,1,N');
       writeln( F,'^XZ');
       writeln( F,'^FX End of job');
       writeln( F,'^XA');
       writeln( F,'^XZ');
       CloseFile(F);
end;


procedure TForm1.StationRadioGroupClick(Sender: TObject);
begin



 if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Triage' then

        begin
              Button1.Caption := 'Enter';
              CbLock.Checked := false;
              LblEmp.Visible := true;
              EdtEmpoyee.Enabled := true;
              EdtEmpoyee.visible := true;
              LblOrderType.visible := true;
              EdtOrderType.visible := true;
              PnlOrdMethod.visible := false;
              LblOrderNo.Visible := true;
              EdtOrder.Visible := true;
              LblContainers.visible := false;
              EdtContainers.visible := false;
              EdtEmpoyee.SetFocus;
              CbLock.Visible := true;
        end;

      if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Assign' then
            begin
              Button1.Caption := 'Enter';
              LblEmp.Visible := true;
              CbLock.Checked := false;
              EdtEmpoyee.Enabled := true;
              EdtEmpoyee.visible := true;
              LblOrderType.visible := false;
              EdtOrderType.visible := false;
              PnlOrdMethod.visible := false;
              LblOrderNo.Visible := true;
              EdtOrder.Visible := true;
              LblContainers.visible := false;
              EdtContainers.visible := false;
              EdtEmpoyee.SetFocus;
              CbLock.Visible := false;
        end;


     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Staging' then
            begin
              Button1.Caption := 'Enter';
              LblEmp.Visible := true;
              CbLock.Checked := false;
              EdtEmpoyee.Enabled := true;
              EdtEmpoyee.visible := true;
              LblOrderType.visible := false;
              EdtOrderType.visible := false;
              PnlOrdMethod.visible := false;
              LblOrderNo.Visible := true;
              EdtOrder.Visible := true;
              LblContainers.visible := false;
              EdtContainers.visible := false;
              EdtEmpoyee.SetFocus;
              CbLock.Visible := false;
        end;

     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'QC' then
           begin
              Button1.Caption := 'Print';
              LblEmp.Visible := true;
              CbLock.Checked := false;
              EdtEmpoyee.Enabled := true;
              EdtEmpoyee.visible := true;
              LblOrderType.visible := false;
              EdtOrderType.visible := false;
              PnlOrdMethod.visible := true;
              LblOrderNo.Visible := true;
              EdtOrder.Visible := true;
              LblContainers.visible := true;
              EdtContainers.visible := true;
              EdtEmpoyee.SetFocus;
              CbLock.Visible := true;
        end;

        if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Dispatch' then
           begin
              Button1.Caption := 'Enter';
              LblEmp.Visible := true;
              CbLock.Checked := false;
              EdtEmpoyee.Enabled := true;
              EdtEmpoyee.visible := true;
              LblOrderType.visible := false;
              EdtOrderType.visible := false;
              PnlOrdMethod.visible := false;
              LblOrderNo.Visible := true;
              EdtOrder.Visible := true;
              LblContainers.visible := false;
              EdtContainers.visible := false;
              EdtEmpoyee.SetFocus;
              CbLock.Visible := false;
        end;
end;


procedure TForm1.Button2Click(Sender: TObject);
var
   Handle: THandle;
   N: DWORD;
   DocInfo1: TDocInfo1;
   sPrinterName : string;
   sTexte : AnsiString;
   sString : AnsiString;
   F: TextFile;
   S: string;
   Description: String;
   Index: Integer;
   PrntName: String;

   BAR_FO_X: String;
   BAR_FO_Y: String;
   BAR_BCN: String;

begin

      BAR_FO_X :=  ('50');
       BAR_FO_Y := ('210');  //240
       BAR_BCN := ('78');

       Description := EdtDescription.text;

      AssignFile(F,'Document1.txt');
      ReWrite(F);

       writeln( F,'^XA^ID*.*^XZ');
       writeln( F,'^XA^PRC');
       writeln( F,'^LH0,0^FS');
       writeln( F,'^LL591');
       writeln( F,'^MD0');
       writeln( F,'^LH0,0^FS');
       writeln( F, '^BY3,3.0^FO' + BAR_FO_X + ',' + BAR_FO_y + '^BCN,' + BAR_BCN + ',Y,N,N^FR^FD>:' + Description + '^FS'); //DEV
       writeln( F,'^PQ1,0,1,N');
       writeln( F,'^XZ');
       writeln( F,'^FX End of job');
       writeln( F,'^XA');
       writeln( F,'^XZ');

      if RGroupPrinters.ItemIndex = 1 then
       begin
        Showmessage('Must be a usb printer');
        exit;       end;

      sPrinterName := Printer.Printers.Strings[Printer.PrinterIndex];

      Caption := sPrinterName;

      if not OpenPrinter(PChar(sPrinterName), Handle, nil) then
      begin
        ShowMessage('error ' + IntToStr(GetLastError));
        Exit;
      end;

      with DocInfo1 do
      begin
        pDocName := PChar('Document1');
        pOutputFile := nil;
        pDataType := 'RAW';
      end;

      Reset(F);
      sString := '';

      while not Eof(F) do
      begin
        ReadLn(F, sTexte);
        sString := sString + sTexte;
      end;

      WinSpool.StartDocPrinter(Handle, 1, @DocInfo1);
      WinSpool.StartPagePrinter(Handle);
      WinSpool.WritePrinter(Handle, PAnsiChar(sString), Length(sString) , N);
      WinSpool.EndPagePrinter(Handle);
      WinSpool.EndDocPrinter(Handle);
      WinSpool.ClosePrinter(Handle);
      CloseFile(F);
end;


procedure TForm1.Button3Click(Sender: TObject);
begin
    EdtEmpoyee.Text := '';
    EdtOrderType.Text := '';
    EdtMethod.Text := '';
    EdtOrder.Text := '';
    EdtContainers.Text := '';
    if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Triage' then
    EdtEmpoyee.SetFocus
    else if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Assign' then
    EdtEmpoyee.SetFocus
    else if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'QC' then
    EdtEmpoyee.SetFocus
    else
    EdtOrder.SetFocus;
end;

procedure TForm1.BtnErrorsClick(Sender: TObject);
var employee, order: String;
var Numeric, Alpha: String;
var i: integer;
var test: string;
var tst: Integer;
begin
screen.Cursor := crHourGlass;

order := EdtErrOrder.text;
employee := EdtErrEmp.Text;

if order = '' then
begin
ShowMessage('Order Number Not Populated');
EdtErrOrder.SetFocus;
screen.Cursor := crDefault;
exit
end;

Alpha := order;
if Alpha[1] in ['a'..'z', 'A'..'Z'] then
begin
  Beep;
  ShowMessage('Incorect Order Number, Try again');
  EdtEmpoyee.text := '';
  EdtErrOrder.SetFocus;
  screen.Cursor := crDefault;
  exit;
end;

if employee = '' then
begin
ShowMessage('Employee Number Not Populated');
EdtErrEmp.SetFocus;
screen.Cursor := crDefault;
exit;
end;

Alpha := employee;
if Alpha[1] in ['a'..'z', 'A'..'Z'] then
begin
  Beep;
  ShowMessage('Incorect Employee Value, Try again');
  EdtEmpoyee.text := '';
  EdtErrEmp.SetFocus;
  screen.Cursor := crDefault;
  exit;
end;


if Length (MemoOther.Text) > 10 then
begin
  etype := 'FreeText';
  code :=  'TEXT'
end;

 if ((ListBoxTriage.ItemIndex = -1) and
   (ListBoxRePrints.ItemIndex = -1) and
   (ListBoxPicking.ItemIndex = -1) and
   (ListBoxPutAway.ItemIndex = -1) and
   (ListBoxPharmacy.ItemIndex = -1) and
   (ListBoxQC.ItemIndex = -1) and
   (ListBoxRetail.ItemIndex = -1) and
   (ListBoxNoErrors.ItemIndex = -1) and
   (MemoOther.Text = '')) then
   begin
    ShowMessage('No Data Entered, Please select error code');
    screen.Cursor := crDefault;
    exit;
   end;

   // Triage

   try
   if not (ListBoxTriage.ItemIndex = -1) then
   begin
      for i := 0 to (ListBoxTriage.Items.Count - 1) do
      begin
        if ListBoxTriage.Checked[i] then
        begin
          case i of
          0: begin code := 'miss'; etype := 'Triage'; end;
          1: begin code := 'mixx'; etype := 'Triage'; end;
          2: begin code := 'ibop'; etype := 'Triage'; end;
          3: begin code := 'nssd'; etype := 'Triage'; end;
          4: begin code := 'wgtry'; etype := 'Triage'; end;
          5: begin code := 'mrprt'; etype := 'Triage'; end;
          6: begin code := 'crprt'; etype := 'Triage'; end;
          else
          Caption := 'None';
          end;
          with ADOInsErrCodes do
          begin
            Close;
            Parameters.ParamByName('ORDERNO').Value := order;
            Parameters.ParamByName('EMP').Value := employee;
            Parameters.ParamByName('TYPE').Value := etype;
            Parameters.ParamByName('CODE').Value := code;
            Parameters.ParamByName('DSC').Value := MemoOther.Text;
            ExecSQL;
          end;
        end;
      end;
   end;
  finally
  Caption := 'Problem Saving Error Code, Please Retry';
end;

// RePrints

try
   if not (ListBoxRePrints.ItemIndex = -1) then
   begin
      for i := 0 to (ListBoxRePrints.Items.Count - 1) do
      begin
        if ListBoxRePrints.Checked[i] then
        begin
          case i of
          0: begin code := 'itemr'; etype := 'Reprint'; end;
          1: begin code := 'itemt'; etype := 'Reprint'; end;
          2: begin code := 'ochng'; etype := 'Reprint'; end;
          3: begin code := 'nnot'; etype := 'Reprint'; end;
          4: begin code := 'typeT'; etype := 'Reprint'; end;
          5: begin code := 'ilot#'; etype := 'Reprint'; end;
          6: begin code := 'iqty'; etype := 'Reprint'; end;
          7: begin code := 'ibin'; etype := 'Reprint'; end;
          8: begin code := 'idts'; etype := 'Reprint'; end;
          9: begin code := 'iaddr'; etype := 'Reprint'; end;
          10: begin code := 'instr'; etype := 'Reprint'; end;
          11: begin code := 'store'; etype := 'Reprint'; end;
          12: begin code := 'dlvy'; etype := 'Reprint'; end;
          13: begin code := 'stock'; etype := 'Reprint'; end;
          else
          Caption := 'None';
          end;
          with ADOInsErrCodes do
          begin
            Close;
            Parameters.ParamByName('ORDERNO').Value := order;
            Parameters.ParamByName('EMP').Value := employee;
            Parameters.ParamByName('TYPE').Value := etype;
            Parameters.ParamByName('CODE').Value := code;
            Parameters.ParamByName('DSC').Value := MemoOther.Text;
            ExecSQL;
          end;
        end;
      end;
   end;
  finally
  Caption := 'Problem Saving Error Code, Please Retry';
end;

// Picking

try
   if not (ListBoxPicking.ItemIndex = -1) then
   begin
      for i := 0 to (ListBoxPicking.Items.Count - 1) do
      begin
        if ListBoxPicking.Checked[i] then
        begin
          case i of
          0: begin code := 'iquan'; etype := 'Picking'; end;
          1: begin code := 'iitem'; etype := 'Picking'; end;
          2: begin code := 'itype'; etype := 'Picking'; end;
          3: begin code := 'ilot'; etype := 'Picking'; end;
          4: begin code := 'istag'; etype := 'Picking'; end;
          5: begin code := 'instr'; etype := 'Picking'; end;
          6: begin code := 'mitem'; etype := 'Picking'; end;
          7: begin code := 'expir'; etype := 'Picking'; end;
          8: begin code := 'mnstk'; etype := 'Picking'; end;
          9: begin code := 'spump'; etype := 'Picking'; end;
          10: begin code := 'initl'; etype := 'Picking'; end;
          11: begin code := 'crrx'; etype := 'Picking'; end;
        else
        Caption := 'None';
        end;
          with ADOInsErrCodes do
          begin
            Close;
            Parameters.ParamByName('ORDERNO').Value := order;
            Parameters.ParamByName('EMP').Value := employee;
            Parameters.ParamByName('TYPE').Value := etype;
            Parameters.ParamByName('CODE').Value := code;
            Parameters.ParamByName('DSC').Value := MemoOther.Text;
            ExecSQL;
          end;
        end;
      end;
   end;
  finally
  Caption := 'Problem Saving Error Code, Please Retry';
end;

// Put Away

try
   if not (ListBoxPutAway.ItemIndex = -1) then
   begin
      for i := 0 to (ListBoxPutAway.Items.Count - 1) do
      begin
        if ListBoxPutAway.Checked[i] then
       begin
      case i of
        0: begin code := 'wbin'; etype := 'PutAway'; end;
        1: begin code := 'mitem'; etype := 'PutAway'; end;
        2: begin code := 'other'; etype := 'PutAway'; end;
      else
      Caption := 'None';
      end;
          with ADOInsErrCodes do
          begin
            Close;
            Parameters.ParamByName('ORDERNO').Value := order;
            Parameters.ParamByName('EMP').Value := employee;
            Parameters.ParamByName('TYPE').Value := etype;
            Parameters.ParamByName('CODE').Value := code;
            Parameters.ParamByName('DSC').Value := MemoOther.Text;
            ExecSQL;
          end;
        end;
      end;
   end;
  finally
  Caption := 'Problem Saving Error Code, Please Retry';
end;

// Pharmacy

try
   if not (ListBoxPharmacy.ItemIndex = -1) then
   begin
      for i := 0 to (ListBoxPharmacy.Items.Count - 1) do
      begin
        if ListBoxPharmacy.Checked[i] then
       begin
        case i of
          0: begin code := 'wday'; etype := 'Pharmacy'; end;
          1: begin code := 'witem'; etype := 'Pharmacy'; end;
          2: begin code := 'mitem'; etype := 'Pharmacy'; end;
          3: begin code := 'miedu'; etype := 'Pharmacy'; end;
          4: begin code := 'perm'; etype := 'Pharmacy'; end;
          5: begin code := 'sdout'; etype := 'Pharmacy'; end;
          else
          Caption := 'None';
        end;
          with ADOInsErrCodes do
          begin
            Close;
            Parameters.ParamByName('ORDERNO').Value := order;
            Parameters.ParamByName('EMP').Value := employee;
            Parameters.ParamByName('TYPE').Value := etype;
            Parameters.ParamByName('CODE').Value := code;
            Parameters.ParamByName('DSC').Value := MemoOther.Text;
            ExecSQL;
          end;
        end;
      end;
   end;
  finally
  Caption := 'Problem Saving Error Code, Please Retry';
end;

// QC

try
   if not (ListBoxQC.ItemIndex = -1) then
   begin
      for i := 0 to (ListBoxQC.Items.Count - 1) do
      begin
        if ListBoxQC.Checked[i] then
       begin
        case i of
          0: begin code := 'iitem'; etype := 'QC/Pack'; end;
          1: begin code := 'itype'; etype := 'QC/Pack'; end;
          2: begin code := 'ilott'; etype := 'QC/Pack'; end;
          3: begin code := 'iqty'; etype := 'QC/Pack'; end;
          4: begin code := 'minst'; etype := 'QC/Pack'; end;
          5: begin code := 'mitem'; etype := 'QC/Pack'; end;
          6: begin code := 'witem'; etype := 'QC/Pack'; end;
          7: begin code := 'expir'; etype := 'QC/Pack'; end;
          8: begin code := 'misns'; etype := 'QC/Pack'; end;
          9: begin code := 'mpump'; etype := 'QC/Pack'; end;
          10: begin code := 'initl'; etype := 'QC/Pack'; end;
          11: begin code := 'frig'; etype := 'QC/Pack'; end;
          12: begin code := 'ipack'; etype := 'QC/Pack'; end;
          13: begin code := 'delry'; etype := 'QC/Pack'; end;
          14: begin code := 'icnt'; etype := 'QC/Pack'; end;
          else
          Caption := 'None';
        end;
          with ADOInsErrCodes do
          begin
            Close;
            Parameters.ParamByName('ORDERNO').Value := order;
            Parameters.ParamByName('EMP').Value := employee;
            Parameters.ParamByName('TYPE').Value := etype;
            Parameters.ParamByName('CODE').Value := code;
            Parameters.ParamByName('DSC').Value := MemoOther.Text;
            ExecSQL;
          end;
        end;
      end;
   end;
  finally
  Caption := 'Problem Saving Error Code, Please Retry';
end;

// Retail
try
   if not (ListBoxRetail.ItemIndex = -1) then
   begin
      for i := 0 to (ListBoxRetail.Items.Count - 1) do
      begin
        if ListBoxRetail.Checked[i] then
       begin
        case i of
          0: begin code := 'ibin'; etype := 'Retail'; end;
          1: begin code := 'iprct'; etype := 'Retail'; end;
          2: begin code := 'inumb'; etype := 'Retail'; end;
          3: begin code := 'patpf'; etype := 'Retail'; end;
          4: begin code := 'sdout'; etype := 'Retail'; end;
          5: begin code := 'sdshp'; etype := 'Retail'; end;
          6: begin code := 'rejct'; etype := 'Retail'; end;
          else
          Caption := 'None';
        end;
          with ADOInsErrCodes do
          begin
            Close;
            Parameters.ParamByName('ORDERNO').Value := order;
            Parameters.ParamByName('EMP').Value := employee;
            Parameters.ParamByName('TYPE').Value := etype;
            Parameters.ParamByName('CODE').Value := code;
            Parameters.ParamByName('DSC').Value := MemoOther.Text;
            ExecSQL;
          end;
        end;
      end;
   end;
  finally
  Caption := 'Problem Saving Error Code, Please Retry';
end;


// No Errors
try
   if not (ListBoxNoErrors.ItemIndex = -1) then
   begin
      for i := 0 to (ListBoxNoErrors.Items.Count - 1) do
      begin
        if ListBoxNoErrors.Checked[i] then
       begin
        case i of
          0: begin code := 'noerr'; etype := 'NoErrors'; end;
          else
          Caption := 'None';
        end;
          with ADOInsErrCodes do
          begin
            Close;
            Parameters.ParamByName('ORDERNO').Value := order;
            Parameters.ParamByName('EMP').Value := employee;
            Parameters.ParamByName('TYPE').Value := etype;
            Parameters.ParamByName('CODE').Value := code;
            Parameters.ParamByName('DSC').Value := MemoOther.Text;
            ExecSQL;
          end;
        end;
      end;
   end;
  finally
  Caption := 'Problem Saving Error Code, Please Retry';
end;


 if ((ListBoxTriage.ItemIndex = -1) and
   (ListBoxRePrints.ItemIndex = -1) and
   (ListBoxPicking.ItemIndex = -1) and
   (ListBoxPutAway.ItemIndex = -1) and
   (ListBoxPharmacy.ItemIndex = -1) and
   (ListBoxQC.ItemIndex = -1) and
   (ListBoxRetail.ItemIndex = -1) and
   (ListBoxNoErrors.ItemIndex = -1) and
   (length(MemoOther.Text) > 3)) then
   begin
      code := 'onlyT'; etype := 'TextOnly';
      with ADOInsErrCodes do
        begin
            Close;
            Parameters.ParamByName('ORDERNO').Value := order;
            Parameters.ParamByName('EMP').Value := employee;
            Parameters.ParamByName('TYPE').Value := etype;
            Parameters.ParamByName('CODE').Value := code;
            Parameters.ParamByName('DSC').Value := MemoOther.Text;
            ExecSQL;
          end;
   end;

  Showmessage('Processed!');
//  Application.MessageBox('Saved','Whs',MB_OK + MB_DEFBUTTON1);
  ErrorsSaved := true;
  BtnClearErrorsClick(Sender);
  Caption := 'Error Code Saved';
  screen.Cursor := crDefault;
  if (overide=false) then
  begin
    PageControl1.ActivePage.PageControl.ActivePageIndex:= 0;
    if CbLock.Checked then
      EdtOrderType.SetFocus
    else
      EdtEmpoyee.SetFocus;
  end;
end;

procedure TForm1.CbLockClick(Sender: TObject);
begin
  if CbLock.Checked then
  begin
     if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Triage' then
     begin
        EdtEmpoyee.Enabled := false;
        EdtEmpoyee.Color := clMenu;
        EdtOrderType.SetFocus;
     end
     else
     begin
        EdtEmpoyee.Enabled := false;
        EdtEmpoyee.Color := clMenu;
        EdtMethod.SetFocus;
     end;
  end
    else
    begin
      EdtEmpoyee.Enabled := true;
      EdtEmpoyee.Color := clWindow;
      EdtEmpoyee.SetFocus;
    end
end;

procedure TForm1.FormShow(Sender: TObject);
begin
    StationRadioGroup.ItemIndex := 0;
    EdtEmpoyee.SetFocus;
    LblProcessed.Caption := '';
end;


procedure TForm1.EdtOrderKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
  var OrderValue : String;
  var Alpha: String;
begin
  if key = VK_RETURN then
  begin
    Alpha := EdtOrder.Text;
    if Alpha[1] in ['a'..'z', 'A'..'Z'] then
    begin
      Beep;
      ShowMessage('Incorect value, Try again');
      EdtOrder.text := '';
      exit;
    end;

    if length(EdtOrder.Text) < 7 then
    begin
      Beep;
      ShowMessage('Incorrect Value Entered!!!.');
      EdtOrder.Text := '';
      EdtOrder.SetFocus;
      exit;
    end;

    if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Triage' then
    begin
     with ADODupChk do
     begin
        Close;
        Parameters.ParamByName('orderno').Value := EdtOrder.Text;
        Open;
     end;
     if  (ADODupChkCNT.AsInteger >= 1) then
     begin
          ShowMessage('Duplicate Order!!!.');
          PageControl1.ActivePage.PageControl.ActivePageIndex:= 5;
          EdtErrOrder.Text := EdtOrder.Text;
          EdtErrEmp.Text := EdtEmpoyee.Text;
          if overide then ErrorsSaved := true else ErrorsSaved := false;
        end
        else
        begin
            Button1Click(Sender);
            if CbLock.Checked then
            EdtOrderType.SetFocus
            else
            EdtEmpoyee.SetFocus;
            exit;
        end;
    end;

   if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Assign' then
   Button1Click(Sender);
   if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Staging' then
   begin
     Button1Click(Sender);
   end;

   if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Dispatch' then
   begin
     Button1Click(Sender);
   end;

   if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'QC' then
   begin
      OrderValue := EdtOrder.text;
      Case OrderValue[1] of
      'U' : begin ShowMessage('Incorrect Value !!!'); EdtOrder.SetFocus; EdtOrder.text := ''; end;
      'D' : begin ShowMessage('Incorrect Value !!!'); EdtOrder.SetFocus; EdtOrder.text := ''; end;
      'C' : begin ShowMessage('Incorrect Value !!!'); EdtOrder.SetFocus; EdtOrder.text := ''; end;
      else
        EdtContainers.SetFocus;
      end;
   end;
  end;
end;

procedure TForm1.EdtOrderTypeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
  var Numeric: String;
begin
  if key = VK_RETURN then
  begin
    Numeric := EdtOrderType.Text;
    if Numeric[1] in ['0'..'9'] then
    begin
      Beep;
      ShowMessage('Incorect value, Try again');
      EdtOrderType.text := '';
      exit;
    end;

    if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Triage' then
    begin
      EdtOrder.SetFocus;
      LblProcessed.Caption := '';
    end
   else
   begin
     ShowMessage('Select correct Order Type !!!');
     Exit;
   end;
  end;
end;

procedure TForm1.EdtMethodKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
  var OrderType: String;
begin
   if key = VK_RETURN then
   begin
      OrderType := EdtMethod.text;
      Case OrderType[1] of
      'U' : EdtOrder.SetFocus;
      'D' : EdtOrder.SetFocus;
      'C' : EdtOrder.SetFocus;
      else
      begin
        ShowMessage('Incorrect Value !!!');
        EdtMethod.text := '';
        EdtMethod.SetFocus;
      end;
      end;
   end;
end;

procedure TForm1.EdtContainersKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = VK_RETURN then
  begin
    if (StrToInt(EdtContainers.Text)) > 30 then
    begin
     ShowMessage('Exceeded Maximum Count !!!');
     EdtContainers.Text := '';
     EdtContainers.SetFocus;
    end
    else
      begin
        LblProcessed.Caption := '';
        Button1Click(Sender);
      end;
  end;
end;

procedure TForm1.EdtEmpoyeeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
  var Numeric, Alpha: String;
begin
   if key = VK_RETURN then
  begin
    Alpha := EdtEmpoyee.Text;
    if Alpha[1] in ['a'..'z', 'A'..'Z'] then
    begin
      Beep;
      ShowMessage('Incorect value, Try again');
      EdtEmpoyee.text := '';
      exit;
    end;

    if length(EdtEmpoyee.Text) > 6 then
    begin
      Beep;
      ShowMessage('Incorrect Value Entered!!!.');
      EdtEmpoyee.Text := '';
      EdtEmpoyee.SetFocus;
      exit;
    end;


//    DBGridDup.Visible := false;
    if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Triage' then
      begin
      EdtOrderType.SetFocus;
      LblProcessed.Caption := '';
      end
    else if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Assign' then
      begin
      EdtOrder.SetFocus;
      LblProcessed.Caption := '';
      end
    else if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Staging' then
      begin
      EdtOrder.SetFocus;
      LblProcessed.Caption := '';
      end
    else if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'QC' then
     begin
      EdtMethod.SetFocus;
      LblProcessed.Caption := '';
      end
    else if StationRadioGroup.Items[StationRadioGroup.ItemIndex] = 'Dispatch' then
      begin
      EdtOrder.SetFocus;
      LblProcessed.Caption := '';
      end
   else
   begin
     ShowMessage('Select correct Order Type !!!');
     Exit;
   end;
  end;
end;

procedure TForm1.EdtErrEmpKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
     if key = VK_RETURN then
      begin
        if EdtErrEmp.Text = '100082' then
        begin
          overide := true;
          ErrorsSaved := true;
          Label12.font.Color := clRed;
        end
        else
        begin
         overide := false;
         ErrorsSaved := false;
         Label12.font.Color := clBlack;
        end;
        MemoOther.SetFocus;
      end;
end;

procedure TForm1.EdtErrOrderKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
       if key = VK_RETURN then
        EdtErrEmp.SetFocus;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  PageControl1.ActivePage.PageControl.ActivePageIndex:= 0;
  EdtOrder.Text := '320718';
  EdtContainers.Text := '';
  EdtMethod.Text := 'Delivery';
  KeepType := False;
  CbLock.Checked := false;
  DBGridDup.Visible := true;
  DateStart.DateTime := date;
  RGroupPrinters.ItemIndex := 0;
  ErrorsSaved := true;
  MemoOther.Clear;
  DbGridErrors.Visible := false;
  overide := false;
end;

end.



