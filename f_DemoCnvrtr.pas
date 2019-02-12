unit f_DemoCnvrtr;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons,
  u_Convert, ComCtrls, ExtCtrls, Gauges, Menus, jpeg;

type
  Td_Frm_Main = class(TForm)
    d_BtBtn_Cnvrt: TBitBtn;
    d_BtBtn_Cls: TBitBtn;
    d_OpnDlg1: TOpenDialog;
    d_SvDlg1: TSaveDialog;
    d_Edt_Opn: TEdit;
    StaticText1: TStaticText;
    d_EdtSv: TEdit;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    Label1: TLabel;
    Label2: TLabel;
    Panel1: TPanel;
    Bevel1: TBevel;
    Gauge2: TGauge;
    Gauge1: TGauge;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N5: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    Image1: TImage;
    Bevel2: TBevel;
    MonthCalendar1: TMonthCalendar;
    N10: TMenuItem;
    N11: TMenuItem;
    d_ChckBx_Diapzn: TCheckBox;
    Label6: TLabel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    d_PppMn_Bttn: TPopupMenu;
    N12: TMenuItem;
    N13: TMenuItem;
    procedure SpeedButton1Click(Sender: TObject);
    procedure d_BtBtn_ClsClick(Sender: TObject);
    procedure d_BtBtn_CnvrtClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure d_ChckBx_DiapznClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure OpenZDir;
  end;

var
  d_Frm_Main: Td_Frm_Main;

implementation

uses f_About, IniFiles;

{$R *.dfm}

procedure Td_Frm_Main.d_BtBtn_ClsClick(Sender: TObject);
begin
   Close;
end;{d_BtBtn_ClsClick}

procedure Td_Frm_Main.d_BtBtn_CnvrtClick(Sender: TObject);
var
   i: integer;
begin
   Label2.Caption:= ' ';
   Label2.Refresh;
   Label2.Caption:= 'Подключение сетевой папки ...';
   Label2.Refresh;
   s_DirNm_Inp:= d_Edt_Opn.Text;
   if s_DirNm_Inp[Length(s_DirNm_Inp)]='\' then Delete(s_DirNm_Inp, Length(s_DirNm_Inp), 1);
   s_DirNm_Out:= d_EdtSv.Text;
   if s_DirNm_Out[Length(s_DirNm_Out)]='\' then Delete(s_DirNm_Out, Length(s_DirNm_Out), 1);
   if Not DirectoryExists(s_DirNm_Out) then
      if not CreateDir(s_DirNm_Out) then
      raise Exception.Create('Не возможно создать папку '+ s_DirNm_Out);

    Gauge1.Progress:= 0;
    Gauge2.Progress:= 0;
    Gauge1.Progress:= 2;
    Gauge2.MaxValue:= MaxFileCount;
    Label2.Caption:= '';
    Label2.Refresh;
    Label2.Caption:= 'Идёт процесс конвертации ...                 ';
    Label2.Refresh;
    i:= ConvertDir;
    Label2.Font.Color:= clBlue;
    Label2.Refresh;
    Label2.Caption:= 'Сконвертировано '+IntToStr(i)+' файл(ов).';
    if i=0 then
    begin
       Gauge1.Progress:= 0;
       Gauge2.Progress:= 0;
       MessageDlg('Файлов для данного конвертирования не существует', mtInformation, [mbYes], 0);
    end;
end;

procedure Td_Frm_Main.FormShow(Sender: TObject);
begin
   Gauge1.Progress:= bPrgrssItm;
end;

procedure Td_Frm_Main.OpenZDir;
begin
   d_OpnDlg1.Filter:= '';
   d_OpnDlg1.DefaultExt:= '';
   d_OpnDlg1.FileName:= {GetCurrentDir} 'Текущая папка';
   if d_OpnDlg1.Execute
   then begin
           s_DirNm_Inp:= GetCurrentDir;
           d_Edt_Opn.Text:= s_DirNm_Inp;
        end else;
end;{OpenZDir}

procedure Td_Frm_Main.SpeedButton1Click(Sender: TObject);
begin
   Gauge1.Progress:= 0;
   Gauge2.Progress:= 0;
   OpenZDir;
   d_OpnDlg1.InitialDir:= GetCurrentDir;
end;

procedure Td_Frm_Main.SpeedButton2Click(Sender: TObject);
begin
   Gauge1.Progress:= 0;
   Gauge2.Progress:= 0;
    d_SvDlg1.Filter:= '';
    d_SvDlg1.DefaultExt:= '';
    d_SvDlg1.FileName:= {GetCurrentDir} 'Текущая папка';
    if d_SvDlg1.Execute
    then begin
            s_DirNm_Out:= GetCurrentDir;
            d_EdtSv.Text:= s_DirNm_Out;
         end;
   d_SvDlg1.InitialDir:= GetCurrentDir;
end;

procedure Td_Frm_Main.N5Click(Sender: TObject);
begin
   OpenZDir;
end;

procedure Td_Frm_Main.N9Click(Sender: TObject);
begin
   d_Frm_About:= Td_Frm_About.Create(Self);
   d_Frm_About.Visible:= True;
end;

procedure Td_Frm_Main.N8Click(Sender: TObject);
begin
   Close;
end;

procedure Td_Frm_Main.N11Click(Sender: TObject);
begin
   if N11.Checked then
   begin
      Image1.Visible:= False;
      Image1.Enabled:= False;
      Bevel2.Visible:= False;
      Bevel2.Enabled:= False;
      MonthCalendar1.Visible:= False;
      MonthCalendar1.Enabled:= False;
      d_Frm_Main.ClientWidth:= 387;
      N11.Checked:= False;
   end
   else
   begin
      Image1.Visible:= True;
      Image1.Enabled:= True;
      Bevel2.Visible:= True;
      Bevel2.Enabled:= True;
      MonthCalendar1.Visible:= True;
      MonthCalendar1.Enabled:= True;
      d_Frm_Main.ClientWidth:= 560;
      N11.Checked:= True;
   end;
end;

procedure Td_Frm_Main.DateTimePicker1Change(Sender: TObject);
begin
   if DateTimePicker1.Date > DateTimePicker2.Date then
   DateTimePicker2.Date:= DateTimePicker1.Date;
end;

procedure Td_Frm_Main.DateTimePicker2Change(Sender: TObject);
begin
   if DateTimePicker2.Date < DateTimePicker1.Date then
   DateTimePicker1.Date := DateTimePicker2.Date;
end;

procedure Td_Frm_Main.d_ChckBx_DiapznClick(Sender: TObject);
begin
   if d_ChckBx_Diapzn.Checked = False
   then begin
           Label6.Enabled:= False;
           DateTimePicker1.Enabled:= False;
           DateTimePicker2.Enabled:= False;
        end
   else begin
           Label6.Enabled:= True;
           DateTimePicker1.Enabled:= True;
           DateTimePicker2.Enabled:= True;
        end;
end;

procedure Td_Frm_Main.FormCreate(Sender: TObject);
var
  Ini: TIniFile;
begin
   Ini:= TIniFile.Create( ChangeFileExt( Application.ExeName, '.dir'));
   try
    s_DirNm_Inp:= Ini.ReadString( 'Dir', 'Inp', 'Z:\HOST\Z_out\ZZZ');
    s_DirNm_Out:= Ini.ReadString( 'Dir', 'Out', 'Z:\HOST\Z_out\ZZZ\Vypiski');
   finally
      Ini.Free;
   end;


   sTempDirOut:= ExtractFilePath(Application.ExeName)+ scTempDir;
   if not DirectoryExists(sTempDirOut) then
      if not CreateDir(sTempDirOut) then
      raise Exception.Create('Не возможно создать папку '+ sTempDirOut);

   Label2.Caption:= ' Введите дату конвертируемых файлов';
   DateTimePicker1.Date:= Date;
   DateTimePicker2.Date:= Date;
   d_Edt_Opn.Text:= s_DirNm_Inp;
   d_EdtSv.Text:= s_DirNm_Out;
   MonthCalendar1.Date:= Date;
end;

procedure Td_Frm_Main.N13Click(Sender: TObject);
begin
   if N13.Checked then
   begin
      SpeedButton1.Enabled:= False;
      SpeedButton2.Enabled:= False;
      N13.Checked:= False;
   end
   else
   begin
      SpeedButton1.Enabled:= True;
      SpeedButton2.Enabled:= True;
      N13.Checked:= True;
   end;
end;

procedure Td_Frm_Main.N12Click(Sender: TObject);
begin
   SpeedButton1.Enabled:= False;
   SpeedButton2.Enabled:= False;
   N13.Checked:= False;
end;

procedure Td_Frm_Main.N7Click(Sender: TObject);
begin
   d_Frm_Main.SpeedButton2Click(Self);
end;

procedure Td_Frm_Main.FormClose(Sender: TObject; var Action: TCloseAction);
var
   Ini: TIniFile;
begin
   Ini := TIniFile.Create( ChangeFileExt( Application.ExeName, '.dir'));
   try
    Ini.WriteString( 'Dir', 'Inp', s_DirNm_Inp);
    Ini.WriteString( 'Dir', 'Out', s_DirNm_Out);
   finally
      Ini.Free;
   end;
end;

end.




