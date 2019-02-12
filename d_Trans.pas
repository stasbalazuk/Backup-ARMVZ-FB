unit d_Trans;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls,
  DateUtils, IniFiles, Gauges, Menus, XPMan, ComCtrls,
  StrUtils,
  Tlhelp32,
  IBDataBase,IBDatabaseInfo,
  CoolTrayIcon, IBServices, IB_Services, pFIBErrorHandler;

type
  Td_Form = class(TForm)
    Panel1: TPanel;
    d_Memo_Info: TMemo;
    Panel2: TPanel;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Label1: TLabel;
    TrayIcon1: TCoolTrayIcon;
    tmr1: TTimer;
    tmr2: TTimer;
    pm1: TPopupMenu;
    Exit1: TMenuItem;
    grp1: TGroupBox;
    d_Lbl_Info: TLabel;
    d_Lbl_Info2: TLabel;
    d_Shape_Led: TShape;
    Shape1: TShape;
    d_Shape_Led2: TShape;
    Shape3: TShape;
    d_BitBtn_Pusk: TBitBtn;
    fibBackUp: TpFIBBackupService;
    fibRestore: TpFIBRestoreService;
    ProgressBar1: TProgressBar;
    procedure OnCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure d_BitBtn_PuskClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure TrayIcon1Click(Sender: TObject);
    procedure tmr1Timer(Sender: TObject);
    procedure tmr2Timer(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    function CopyFile(sOldFile, sNewFile: string): boolean;
    function ProgressMax: longint;
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
    SessionEnding: Boolean;
    function LedSwich: boolean;
    function LedSwich2: boolean;
    procedure WMQueryEndSession(var Message: TMessage); message WM_QUERYENDSESSION;
  public
    { Public declarations }
  end;

type
   TRConfigIniKhD = record
      sMask: string;
      sPathLogLocal: string;
      sPathLogServer: string;
      sPathNetServer: string;
      sPathRab: string;
      sPathBackupServer: string;
      sPathBackupLocal: string;
      dPreDate: TDate;
      dLastDate: TDate;
      iCountMoon: integer;
    end;

var
   d_Form: Td_Form;
   backid: string;
   recConfigIni: TRConfigIniKhD;
   bFlConnect: boolean;
   iCount: Longint = 0;
   iCntMn: longint = 0;
   sLogFileName: string;
   mMask,bday,bcday: string;
   R: TStringList;
   s,s1,s2,pth,ips: string;

implementation

{$R *.dfm}

procedure Td_Form.WMQueryEndSession(var Message: TMessage);
begin
  SessionEnding := True;
  Message.Result := 1;
  Application.Terminate;
end;

procedure CreateFormInRightBottomCorner;
var
 r : TRect;
begin
 SystemParametersInfo(SPI_GETWORKAREA, 0, Addr(r), 0);
 d_Form.Left := r.Right-d_Form.Width;
 d_Form.Top := r.Bottom-d_Form.Height;
end;

function KillTask(ExeFileName: string): Integer;
const
  PROCESS_TERMINATE = $0001;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
  Result := 0;
  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := SizeOf(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);

  while Integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) =
      UpperCase(ExeFileName)) or (UpperCase(FProcessEntry32.szExeFile) =
      UpperCase(ExeFileName))) then
      Result := Integer(TerminateProcess(
                        OpenProcess(PROCESS_TERMINATE,
                                    BOOL(0),
                                    FProcessEntry32.th32ProcessID),
                                    0));
     ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
  end;
  CloseHandle(FSnapshotHandle);
end;

// запись в реестра
function RegWriteStr(RootKey: HKEY; Key, Name, Value: string): Boolean;
var
  Handle: HKEY;
  Res: LongInt;
begin
  Result := False;
  Res := RegCreateKeyEx(RootKey, PChar(Key), 0, nil, REG_OPTION_NON_VOLATILE,
    KEY_ALL_ACCESS, nil, Handle, nil);
  if Res <> ERROR_SUCCESS then
    Exit;
  Res := RegSetValueEx(Handle, PChar(Name), 0, REG_SZ, PChar(Value),
    Length(Value) + 1);
  Result := Res = ERROR_SUCCESS;
  RegCloseKey(Handle);
end; 

function SplitStr(s: string): string;
begin
   Result:= s;
   if s = '' then Exit;
   if s[Length(s)]<>'\' then Result:= s+'\';
end;{SplitStr}

function StrTime: string;
begin
   Result:= TimeToStr(GetTime) +'  ';
end;{StrTime}

// CopyFileEx
function CopyCallBack(
  TotalFileSize: LARGE_INTEGER;          // Taille totale du fichier en octets
  TotalBytesTransferred: LARGE_INTEGER;  // Nombre d'octets dйjаs transfйrйs
  StreamSize: LARGE_INTEGER;             // Taille totale du flux en cours
  StreamBytesTransferred: LARGE_INTEGER; // Nombre d'octets dйjа tranfйrйs dans ce flus
  dwStreamNumber: DWord;                 // Numйro de flux actuem
  dwCallbackReason: DWord;               // Raison de l'appel de cette fonction
  hSourceFile: THandle;                  // handle du fichier source
  hDestinationFile: THandle;             // handle du fichier destination
  ProgressBar : TProgressBar             // paramиtre passй а la fonction qui est une
                                         // recopie du paramиtre passй а CopyFile Ex
                                         // Il sert а passer l'adresse du progress bar а
                                         // mettre а jour pour la copie. C'est une
                                         // excellente idйe de DelphiProg
  ): DWord; far; stdcall;
var
  EnCours: Int64;
begin
  EnCours := TotalBytesTransferred.QuadPart * 100 div TotalFileSize.QuadPart;
  If ProgressBar<>Nil Then ProgressBar.Position := EnCours;
     Result := PROGRESS_CONTINUE;
end;

function Td_Form.LedSwich: boolean;
begin
   if DirectoryExists(recConfigIni.sPathLogLocal)
   then begin
         d_Lbl_Info.Font.Color:= clGreen;
         d_Lbl_Info.Caption:= 'Связь с сервером АРМВЗ - ОК';
         bFlConnect:= True;
         d_Shape_Led.Brush.Color:= clLime;
         d_Shape_Led.Pen.Color:= clGreen;
         Result:= True;
        end
   else begin
         d_Lbl_Info.Font.Color:= clRed;
         d_Lbl_Info.Caption:= 'Нет связи с сервером АРМВЗ';
         bFlConnect:= False;
         d_Shape_Led.Brush.Color:= clRed;
         d_Shape_Led.Pen.Color:= clMaroon;
         Result:= False;
        end;
end;{LedSwich}

function Td_Form.LedSwich2: boolean;
begin
   if DirectoryExists(recConfigIni.sPathBackupServer)
   then begin
         d_Lbl_Info2.Font.Color:= clGreen;
         d_Lbl_Info2.Caption:= 'Связь с сервером ЦПЗ - ОК';
         bFlConnect:= True;
         d_Shape_Led2.Brush.Color:= clLime;
         d_Shape_Led2.Pen.Color:= clGreen;
         Result:= True;
        end
   else begin
         d_Lbl_Info2.Font.Color:= clRed;
         d_Lbl_Info2.Caption:= 'Нет связи с сервером ЦПЗ';
         bFlConnect:= False;
         d_Shape_Led2.Brush.Color:= clRed;
         d_Shape_Led2.Pen.Color:= clMaroon;
         Result:= False;
        end;
end;{LedSwich}

function LogDat: string;
var
   sDat: string;
begin
   sDat:= DateToStr(Date);
   Result:= 'Log_'+ sDat[1]+sDat[2]+'-'+sDat[4]+sDat[5]+'-'+sDat[9]+sDat[10];
end;{LogDat}

procedure Td_Form.OnCreate(Sender: TObject);
var
  Ini: TIniFile;
  iniF1: string;
  i: integer;
begin
   backid:='0';
   CreateFormInRightBottomCorner;
   RegWriteStr(HKEY_CURRENT_USER,'SOFTWARE\Microsoft\Windows\CurrentVersion\Run','Backup_GDB',ParamStr(0));
   iniF1:=ExtractFilePath(ParamStr(0))+'TransFiles.ini';  //TransFiles.ini
   d_BitBtn_Pusk.Tag:=10;
   d_BitBtn_Pusk.Enabled:=False;
   if FileExists(iniF1) then begin
   Ini:= TIniFile.Create(iniF1);
   try
    recConfigIni.sPathLogLocal:= SplitStr(Ini.ReadString('LocalDIR','LocalDIR',''));
    recConfigIni.sPathLogServer:= SplitStr(Ini.ReadString('LocalDIR','LocalDIRBackup',''));
    R:=TStringList.Create;
    ExtractStrings(['\'],['\'],PChar(recConfigIni.sPathLogLocal),R);
    for i:=0 to R.Count-1 do begin
        s1:=R.Strings[i];
        if i = 1 then
           s:=s+s1+':\'
        else if i = 0 then begin
           s:=s1+':';
           ips:=s1;
        end else s:=s+s1+'\';
        s:=AnsiReplaceStr(s, ' ', '');
    end;
    pth:=AnsiReplaceStr(s, ' ', '');
    R.Free;
    recConfigIni.sMask:= Ini.ReadString('ToCPZ','Mask','');
    recConfigIni.sPathBackupServer:= SplitStr(Ini.ReadString('ToCPZ','DestDIR',''));
    recConfigIni.sPathBackupLocal:= SplitStr(Ini.ReadString('ToCPZ','DestDIRBackup',''));
    backid:= SplitStr(Ini.ReadString('Backup','Count',''));
    mMask:=recConfigIni.sMask+Copy(DateToStr(Now),7,4)+Copy(DateToStr(Now),4,2)+'.GDB'; //REG201603.GDB
   finally
      Ini.Free;
   end;
   end;
   d_Memo_Info.Lines.Add('Старт программы '+ DateToStr(Date)+' в '+ TimeToStr(GetTime));
   d_Memo_Info.Lines.Add(StrTime+Caption);
   d_Memo_Info.Lines.Add(StrTime+ 'Проверка связи с серверами ...');
   d_Memo_Info.Lines.Add(StrTime+ '==============================');
   iCount:= 0;
   bday:=DateToStr(Date-1);
   bcday:=DateToStr(Date-2);
   if LedSwich
   then begin
        bFlConnect:= True;
          if not DirectoryExists(recConfigIni.sPathLogServer) then begin
             if not ForceDirectories(recConfigIni.sPathLogServer) then
             begin
               raise Exception.Create('Невозможно создать каталог '+ recConfigIni.sPathLogServer);
               d_Memo_Info.Lines.Add(StrTime+ recConfigIni.sPathLogServer);
               bFlConnect:= False;
             end;
          end else begin
               d_BitBtn_Pusk.Tag:=0;
               d_BitBtn_Pusk.Enabled:=True;
               d_Memo_Info.Lines.Add(StrTime+ 'Связь с сервером АРМВЗ - ОК');
          end;
          bday:=DateToStr(Date-1);
          bcday:=DateToStr(Date-2);
          s:=recConfigIni.sPathBackupServer+bday+'\';
          recConfigIni.sPathBackupServer:=s;
          if not DirectoryExists(recConfigIni.sPathBackupServer) then begin
             if not ForceDirectories(recConfigIni.sPathBackupServer) then
             begin
               raise Exception.Create('Невозможно создать каталог '+ recConfigIni.sPathBackupServer);
               d_Memo_Info.Lines.Add(StrTime+ recConfigIni.sPathBackupServer);
               bFlConnect:= False;
             end;
          end else begin
               d_BitBtn_Pusk.Tag:=0;
               d_BitBtn_Pusk.Enabled:=True;
               d_Memo_Info.Lines.Add(StrTime+ 'Связь с сервером АРМВЗ - ОК');
          end;
   end
   else begin
        d_BitBtn_Pusk.Tag:=10;
        d_BitBtn_Pusk.Enabled:=False;
        bFlConnect:= False;
        d_Memo_Info.Lines.Add(StrTime+ 'Нет связи с сервером АРМВЗ: '+recConfigIni.sPathBackupServer);
   end;
   if LedSwich2
   then begin
        bFlConnect:= True;
          if not DirectoryExists(recConfigIni.sPathBackupServer) then begin
             if not ForceDirectories(recConfigIni.sPathBackupServer) then
             begin
               raise Exception.Create('Невозможно создать каталог '+ recConfigIni.sPathBackupServer);
               d_Memo_Info.Lines.Add(StrTime+ recConfigIni.sPathBackupServer);
               bFlConnect:= False;
             end;
          end else begin
               d_BitBtn_Pusk.Tag:=0;
               d_BitBtn_Pusk.Enabled:=True;
               d_Memo_Info.Lines.Add(StrTime+ 'Связь с сервером ЦПЗ - ОК');
          end;
          if not DirectoryExists(recConfigIni.sPathBackupLocal) then begin
             if not ForceDirectories(recConfigIni.sPathBackupLocal) then
             begin
               raise Exception.Create('Невозможно создать каталог '+ recConfigIni.sPathBackupLocal);
               d_Memo_Info.Lines.Add(StrTime+ recConfigIni.sPathBackupLocal);
               bFlConnect:= False;
             end;
          end else begin
               d_BitBtn_Pusk.Tag:=0;
               d_BitBtn_Pusk.Enabled:=True;
               d_Memo_Info.Lines.Add(StrTime+ 'Связь с сервером АРМВЗ - ОК');
          end;
   end
   else begin
        d_BitBtn_Pusk.Tag:=10;
        d_BitBtn_Pusk.Enabled:=False;
        bFlConnect:= False;
        d_Memo_Info.Lines.Add(StrTime+ 'Нет связи с сервером ЦПЗ: '+recConfigIni.sPathBackupServer);
   end;
   d_Memo_Info.Lines.Add(StrTime+ '==============================');
end;{Td_Form.OnCreate}

procedure Td_Form.FormClose(Sender: TObject; var Action: TCloseAction);
var
   Ini: TIniFile;
begin
  if not FileExists(ChangeFileExt(Application.ExeName,'.ini')) then begin
    Ini := TIniFile.Create( ChangeFileExt( Application.ExeName, '.ini'));
   try
    Ini.WriteString('LocalDIR','LocalDIR',recConfigIni.sPathLogLocal);
    Ini.WriteString('LocalDIR','LocalDIRBackup',recConfigIni.sPathLogServer);
    Ini.WriteString('ToCPZ','Mask1',recConfigIni.sMask);
    Ini.WriteString('ToCPZ','DestDIR',recConfigIni.sPathBackupServer);
    Ini.WriteString('ToCPZ','DestDIRBackup',recConfigIni.sPathBackupLocal);
    Ini.WriteString('Date','PreDate',DateToStr(Now));
    Ini.WriteString('Date','LastDate',DateToStr(Now));
    Ini.WriteString('Date','CountMoon',IntToStr(MonthOfTheYear(Now)));
    Ini.WriteString('Backup','Count',backid);
   finally
      Ini.Free;
   end;
  end;
end;

function Td_Form.CopyFile(sOldFile, sNewFile: string): boolean;
var
//  NewFileName, OldFileName: string;
//  Msg: string;
  NewFile: TFileStream;
  OldFile: TFileStream;
begin
//  Msg := Format('Copy %s to %s?', [Edit1.Text, NewFileName]);
//  if MessageDlg(Msg, mtCustom, mbOKCancel, 0) = mrOK then
    Result:= True;
    try
      OldFile := TFileStream.Create(sOldFile, fmOpenRead or fmShareDenyWrite);
      NewFile := TFileStream.Create(sNewFile, fmCreate {or fmShareDenyRead});
      try
         if NewFile.CopyFrom(OldFile, OldFile.Size)<> OldFile.Size
         then Result:= False;
         except
           Result:= False;
         end;
    finally
       FreeAndNil(OldFile);
       FreeAndNil(NewFile)
    end;
end;{CopyFile}


function Td_Form.ProgressMax: longint;
var
   i: longint;
   SR: TSearchRec;
   sFName: TFileName;
begin
   i:=0;
   with recConfigIni do
   try
     if FindFirst(SplitStr(sPathNetServer)+ sMask, faAnyFile, SR)<>0
     then Result:= 0
     else begin
             repeat
                Application.ProcessMessages;
                if LedSwich
                then begin
                        sFName:= SR.Name;
                        Inc(i);
                        bFlConnect:= True
                     end
                else begin
                        bFlConnect:= False;
                        d_Memo_Info.Lines.Add(StrTime + 'Остановлено пользователем');
                     end;
             until ((FindNext(SR)<>0) or not bFlConnect);
             if not bFlConnect then d_Memo_Info.Lines.Add(StrTime + 'Остановлено пользователем');
             Result:= i;
          end;
    finally
       FindClose(sr);
    end;
end;

function ExtractOnlyFileName(const FileName: string): string;
begin
  result:=StringReplace(ExtractFileName(FileName),ExtractFileExt(FileName),'',[]);
end;

procedure Td_Form.d_BitBtn_PuskClick(Sender: TObject);
var
   SR: TSearchRec;
   sFName: TFileName;
   i: Integer;
   StrL : TStringList;
   Retour: LongBool;
   IBDB:TIBDataBase;
   IBInf:TIBDatabaseInfo;
   ConnectedUsers:TStringList;
begin
   if d_BitBtn_Pusk.Tag = 10 then Exit;
   d_BitBtn_Pusk.Enabled:=False;
   Label1.Caption:= '0';
   bFlConnect:= True;
   StrL := TStringList.Create;
   StrL.Delimiter := ',';
   StrL.DelimitedText := mMask;
   KillTask('gbak.exe');
   KillTask('gbak.exe');
   KillTask('gbak.exe');
   for i := 0 to StrL.Count-1 do begin
   mMask:=StrL.Strings[i];
   with recConfigIni do
   try
   if FindFirst(SplitStr(recConfigIni.sPathLogLocal)+ mMask, faAnyFile, SR)=0 then
      sFName:=SR.Name;
      tmr1.Enabled:=False;
      tmr2.Enabled:=False;
     if LedSwich and LedSwich2 then
     if FileExists(recConfigIni.sPathLogLocal+sFName) then begin
        d_Memo_Info.Lines.Add(StrTime+'Копирование файла ... '+sFName);
             repeat
                Application.ProcessMessages;
                if FileExists(recConfigIni.sPathLogLocal+sFName) then begin
                   Retour := False;
                   if DirectoryExists(recConfigIni.sPathLogServer) then
                   if not CopyFileEx(
                      PChar(recConfigIni.sPathLogLocal+sFName),
                      PChar(recConfigIni.sPathLogServer+sFName),
                      @CopyCallBack,
                      ProgressBar1,
                      @Retour,
                      COPY_FILE_RESTARTABLE)
                   then d_Memo_Info.Lines.Add(StrTime + IntToStr(GetLastError));
                   if (LedSwich and LedSwich2)
                    then begin
                      Inc(iCount);
                      d_Memo_Info.Lines.Add(StrTime + recConfigIni.sPathLogServer+sFName + ' - Ok');
                      Application.ProcessMessages;
                    end else begin
                      bFlConnect:= False;
                      d_Memo_Info.Lines.Add(StrTime + 'Нет связи с сервером!! ');
                      Exit;
                    end;
                   if DirectoryExists(recConfigIni.sPathBackupServer) then
                   if not CopyFileEx(
                      PChar(recConfigIni.sPathLogLocal+sFName),
                      PChar(recConfigIni.sPathBackupServer+sFName),
                      @CopyCallBack,
                      ProgressBar1,
                      @Retour,
                      COPY_FILE_RESTARTABLE)
                   then d_Memo_Info.Lines.Add(StrTime + IntToStr(GetLastError));
                   if (LedSwich and LedSwich2)
                    then begin
                      Inc(iCount);
                      d_Memo_Info.Lines.Add(StrTime + recConfigIni.sPathBackupServer+sFName + ' - Ok');
                      Application.ProcessMessages;
                    end else begin
                      bFlConnect:= False;
                      d_Memo_Info.Lines.Add(StrTime + 'Нет связи с сервером!! ');
                    end;
                   if DirectoryExists(recConfigIni.sPathBackupLocal) then
                   if not CopyFileEx(
                      PChar(recConfigIni.sPathLogLocal+sFName),
                      PChar(recConfigIni.sPathBackupLocal+sFName),
                      @CopyCallBack,
                      ProgressBar1,
                      @Retour,
                      COPY_FILE_RESTARTABLE)
                   then d_Memo_Info.Lines.Add(StrTime + IntToStr(GetLastError));
                   if (LedSwich and LedSwich2)
                    then begin
                      Inc(iCount);
                      d_Memo_Info.Lines.Add(StrTime + recConfigIni.sPathBackupLocal+sFName + ' - Ok');
                      Application.ProcessMessages;
                    end else begin
                      bFlConnect:= False;
                      d_Memo_Info.Lines.Add(StrTime + 'Нет связи с сервером!! ');
                    end;
                end;
             until ((FindNext(SR)<>0) or not bFlConnect);
             if not bFlConnect then d_Memo_Info.Lines.Add(StrTime + 'Остановлено пользователем');
             FindClose(sr);
             d_Memo_Info.Lines.Add(StrTime + 'Копирование файлов завершено');
             d_Memo_Info.Lines.Add(StrTime + 'Всего скопировано '+ IntToStr(iCount)+' файла');
             d_Memo_Info.Lines.Add('========================================');
             if FileExists(recConfigIni.sPathBackupServer+sFName) then begin
              try
              IBDB:=TIBDataBase.Create(nil);
              IBDB.DatabaseName:=pth+sFName;
              IBDB.LoginPrompt:=false;
              IBDB.Params.Add('user_name=SYSDBA');
              IBDB.Params.Add('password=masterkey');
              IBDB.Params.Add('lc_ctype=win1251');
              IBDB.Open;
              IBInf:=TIBDatabaseInfo.Create(nil);
              IBInf.Database:=IBDB;
              if IBInf.UserNames.Count > 0 then begin
                  ExitCode:=2;
                  d_Memo_Info.Lines.Add(StrTime + 'ACTIVE CONNECTIONS = '+IntToStr(IBInf.UserNames.Count));
                  ConnectedUsers:=IBInf.UserNames;
                  if ConnectedUsers.IndexOf(ParamStr(2))<>-1 then
                  ConnectedUsers.Delete(ConnectedUsers.IndexOf(ParamStr(2)));
                  d_Memo_Info.Lines.Add(StrTime + 'USER: '+ConnectedUsers.Text);
               if FileExists(recConfigIni.sPathBackupServer+sFName) then begin
                  fibBackUp.Protocol:=TCP;
                  fibBackUp.LoginPrompt:=false;
                  fibBackUp.DatabaseName := pth+sFName;
                  fibBackUp.ServerName := 'localhost';
                  fibBackUp.BackupFile.Add(recConfigIni.sPathBackupServer + ExtractFileName(pth+sFName)+ '_' + DateToStr(now) + '.gbk');
                  fibBackUp.Params.Add('user_name=SYSDBA');
                  fibBackUp.Params.Add('password=masterkey');
                  fibBackUp.Active := True;
                  try
                    Screen.Cursor := crSQLWait;
                    fibBackUp.ServiceStart;
                    d_Memo_Info.Lines.Add(StrTime + '**************** Резервное копирование базы: ' + sFName + '****************');
                    ProgressBar1.Position:=0;
                    ProgressBar1.Max:=fibBackUp.InstanceSize;
                    while not (fibBackUp.Eof) do begin
                       s:=fibBackUp.GetNextLine;
                       s:=Trim(s);
                       d_Memo_Info.Lines.Add(StrTime + s);
                       Application.ProcessMessages;
                    end;
                    d_Memo_Info.Lines.Add(StrTime + '*************** Резервное копирование закончено ***************');
                    d_Memo_Info.Lines.SaveToFile(ExtractFilePath(ParamStr(0))+'Log.txt');
                    fibBackUp.Active := false;
                    Screen.Cursor := crDefault;
                  except
                    MessageDlg('Ошибка! Резервного копирования базы '+ sFName,mtError,[mbOk],0);
                  end;
               end;
              end else begin
                  ExitCode:=0;
                  d_Memo_Info.Lines.Add(StrTime + 'NO ACTIVE CONNECTIONS');
              end;
              except
              on E:Exception do
                begin
                  d_Memo_Info.Lines.Add(StrTime + 'Error='+E.Message);
                  ExitCode:=1;
                end;
              end;
             end;
             TrayIcon1.ShowBalloonHint('Внимание', 'Копирование файлов завершено!', bitInfo, 11);
             iCount:= 0;
             sFName:='';
             tmr1.Enabled:=True;
             tmr2.Enabled:=True;
             KillTask('gbak.exe');
             KillTask('gbak.exe');
             KillTask('gbak.exe');
     end else begin
        d_Memo_Info.Lines.Add(StrTime+'Нет файлов '+mMask+' для копирования!');
     end;
   except
      d_Memo_Info.Lines.Add(StrTime +'Ошибка копирования!');
   end;
   end;
   StrL.Free;
   tmr1.Enabled:=True;
   tmr2.Enabled:=True;
   d_BitBtn_Pusk.Enabled:=True;
   d_Memo_Info.Lines.SaveToFile(ExtractFilePath(ParamStr(0))+'Log.txt');
end;{Td_Form.d_BitBtn_PuskClick}

procedure Td_Form.FormDestroy(Sender: TObject);
begin
  KillTask('gbak.exe');
  KillTask('gbak.exe');
  KillTask('gbak.exe');
  TrayIcon1.IconVisible:=False;
end;

procedure Td_Form.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
    CanClose := SessionEnding;
 if not CanClose then
  begin
    TrayIcon1.HideMainForm;
    TrayIcon1.IconVisible := True;
  end;
end;

procedure Td_Form.TrayIcon1Click(Sender: TObject);
begin
  TrayIcon1.ShowMainForm;
  TrayIcon1.IconVisible:=False;
end;

procedure Td_Form.tmr1Timer(Sender: TObject);
var
  tm,s,s1: string;
begin
  tm:=TimeToStr(now);
  s:=tm[5];
  s1:=tm[4];
  Label1.Caption:='['+s1+s+']';
  if s <> '0' then begin
   tmr1.Enabled:=False;
   d_Memo_Info.Clear;
   d_Memo_Info.Lines.Add('Старт программы '+ DateToStr(Date)+' в '+ TimeToStr(GetTime));
   d_Memo_Info.Lines.Add(StrTime+Caption);
   d_Memo_Info.Lines.Add(StrTime+ 'Проверка связи с серверами ...');
   d_Memo_Info.Lines.Add(StrTime+ '==============================');
   iCount:= 0;
   if LedSwich
   then begin
        bFlConnect:= True;
          if not DirectoryExists(recConfigIni.sPathLogServer) then begin
             if not ForceDirectories(recConfigIni.sPathLogServer) then
             begin
               raise Exception.Create('Невозможно создать каталог '+ recConfigIni.sPathLogServer);
               d_Memo_Info.Lines.Add(StrTime+ recConfigIni.sPathLogServer);
               bFlConnect:= False;
             end;
          end else begin
               d_BitBtn_Pusk.Tag:=0;
               d_BitBtn_Pusk.Enabled:=True;
               d_Memo_Info.Lines.Add(StrTime+ 'Связь с сервером АРМВЗ - ОК');
          end;
   end
   else begin
        d_BitBtn_Pusk.Tag:=10;
        d_BitBtn_Pusk.Enabled:=False;
        bFlConnect:= False;
        d_Memo_Info.Lines.Add(StrTime+ 'Нет связи с сервером АРМВЗ: '+recConfigIni.sPathLogServer);
   end;
   if LedSwich2
   then begin
        bFlConnect:= True;
          if not DirectoryExists(recConfigIni.sPathBackupServer) then begin
             if not ForceDirectories(recConfigIni.sPathBackupServer) then
             begin
               raise Exception.Create('Невозможно создать каталог '+ recConfigIni.sPathBackupServer);
               d_Memo_Info.Lines.Add(StrTime+ recConfigIni.sPathBackupServer);
               bFlConnect:= False;
             end;
          end else begin
               d_BitBtn_Pusk.Tag:=0;
               d_BitBtn_Pusk.Enabled:=True;
               d_Memo_Info.Lines.Add(StrTime+ 'Связь с сервером ЦПЗ - ОК');
          end;
   end
   else begin
        d_BitBtn_Pusk.Tag:=10;
        d_BitBtn_Pusk.Enabled:=False;
        bFlConnect:= False;
        d_Memo_Info.Lines.Add(StrTime+ 'Нет связи с сервером ЦПЗ: '+recConfigIni.sPathBackupServer);
   end;
   d_Memo_Info.Lines.Add(StrTime+ '==============================');
   tmr2.Enabled:=True;
  end;
  if s = '0' then begin
     tmr1.Enabled:=False;
  if tm = '10:10:00' then begin
     tmr1.Enabled:=False;
     d_BitBtn_Pusk.Click;
     tmr1.Enabled:=True;
  end;
  if tm = '12:10:00' then begin
     tmr1.Enabled:=False;
     d_BitBtn_Pusk.Click;
     tmr1.Enabled:=True;
  end;
  if tm = '15:10:00' then begin
     tmr1.Enabled:=False;
     d_BitBtn_Pusk.Click;
     tmr1.Enabled:=True;
  end;
     tmr2.Enabled:=True;
  end;
end;

procedure Td_Form.tmr2Timer(Sender: TObject);
var
  tm,s: string;
begin
  tm:=TimeToStr(now);
  s:=tm[5];
  if s <> '0' then begin
     tmr2.Enabled:=False;
     tmr1.Enabled:=True;
  end;
end;

procedure Td_Form.Exit1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure Td_Form.FormActivate(Sender: TObject);
begin
CreateFormInRightBottomCorner;
end;

end.
