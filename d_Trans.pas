unit d_Trans;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls,
  DateUtils, IniFiles, Gauges, Menus, XPMan, ComCtrls,
  StrUtils,
  ShellAPI,
  Tlhelp32,
  masks,
  IBDataBase,IBDatabaseInfo,
  CoolTrayIcon, IBServices, IB_Services, pFIBErrorHandler, jpeg;

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
    lg0: TGroupBox;
    lst1: TListBox;
    img1: TImage;
    procedure OnCreate(Sender: TObject);
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
    procedure FindFiles(Path, Mask: string; List: TStrings; IncludeSubDir: Boolean = True);
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
      sPathBackupServer: string;
      sFirebird: string;
      sTimeBegin: string;
      sTimeEnd: string;
    end;

const
   bgt = '10:10';
   ent = '10:11';

var
   d_Form: Td_Form;
   errs: Integer;
   backid,fireb: string;
   recConfigIni: TRConfigIniKhD;
   bFlConnect,stime: boolean;
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

//Проверка используется ли формат времени в 24 часа
function Is24HourTimeFormat: Boolean;
var
  DefaultLCID: LCID;
begin
  DefaultLCID := GetThreadLocale;
  Result := 0 <> StrToIntDef(GetLocaleStr(DefaultLCID, LOCALE_ITIME,'0'), 0);
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

function SDStr(s: string): string;
var
  i: integer;
begin
   Result:= s;
   if s = '' then Exit;
   if s[Length(s)]='\' then begin
      i:=Length(s);
      Delete(s,i,i-1);
      Result:= s;
   end;
end;

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
         d_Lbl_Info2.Caption:= 'Связь с сервером Backup - ОК';
         bFlConnect:= True;
         d_Shape_Led2.Brush.Color:= clLime;
         d_Shape_Led2.Pen.Color:= clGreen;
         Result:= True;
        end
   else begin
         d_Lbl_Info2.Font.Color:= clRed;
         d_Lbl_Info2.Caption:= 'Нет связи с сервером Backup';
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

// Функция удаления директории с вложениями.
{
if MyDirectoryDelete('myfolder') then
ShowMessage('Папка успешно удалена.')
else
ShowMessage('Ошибка: папка не удалена.');
}
function MyDirectoryDelete(dir: string): Boolean;
var
fos: TSHFileOpStruct;
begin
ZeroMemory(@fos, SizeOf(fos));
with fos do begin
wFunc := FO_DELETE;
fFlags := FOF_SILENT or FOF_NOCONFIRMATION;
pFrom := PChar(dir + #0);
end;
Result := (0 = ShFileOperation(fos));
end;

function GetFileDate(FileName: string): string;
var
  a,b,c:TFileTime;
  tm:TSystemTime;
  hFile:THandle;
begin
  Result:='';
  hFile:=FileOpen(FileName,fmOpenWrite);
  try
    GetFileTime(hFile,@a,@b,@c);
    FileTimeToSystemTime(a,tm);
    Result:=DateTimeToStr(SystemTimeToDateTime(tm));
  finally
      FileClose(hFile);
  end;
end;

function GFileDate(FileName: string): string;
var
  FHandle: Integer;
begin
  FHandle := FileOpen(FileName, 0);
  try
    Result := DateTimeToStr(FileDateToDateTime(FileGetDate(FHandle)));
  finally
    FileClose(FHandle);
  end;
end;

procedure Td_Form.FindFiles(Path, Mask: string; List: TStrings; IncludeSubDir: Boolean = True);
var
  SearchRec: TSearchRec;
  FindResult: Integer;
  s: string;
  i: Integer;
begin
  d_Form.d_Memo_Info.Update;
try
  Path:=IncludeTrailingBackSlash(Path);
  FindResult:=FindFirst(Path+'*.*', faAnyFile, SearchRec);
  try
    while FindResult = 0 do with SearchRec do begin
        if (Attr and faDirectory<>0) then begin
            if IncludeSubDir and (Name<>'..') and (Name<>'.') then begin
               FindFiles(Path+Name, Mask, List, IncludeSubDir);
            end;
        end else begin
        if MatchesMask(Name, Mask) then begin
              s:=GFileDate(Path+SearchRec.Name);
              i:=Pos(' ',s);
              s:=Trim(Copy(s,1,i));
              if s < DateToStr(Date-3) then begin
                 List.Add(Path);
              end;
        end;
        end;
        FindResult:=FindNext(SearchRec);
    end;
  finally
        FindClose(SearchRec);
  end;
finally
  d_Form.d_Memo_Info.Update;
end;
end;

procedure RemoveEmptyDirectories(rootPath: String);
var
	searchRec: TSearchRec;
begin
	if FindFirst(rootPath + '*', faDirectory, searchRec) = 0 then
	begin
		repeat
			if (searchRec.attr and faDirectory) = faDirectory then
			begin
				if (searchRec.Name <> '.') and (searchRec.Name <> '..') then
				begin
					RemoveDir(rootPath + searchRec.Name);
				end;
			end;
		until FindNext(searchRec) <> 0;

		FindClose(searchRec);
	end;
end;

procedure Td_Form.OnCreate(Sender: TObject);
var
  Ini: TIniFile;
  iniF1: string;
  i,y: integer;
  TST: TStringList;
begin
   backid:='0';
   errs:=0;
   stime:=False;
   TST:= TStringList.Create;
   CreateFormInRightBottomCorner;
   if FileExists(ChangeFileExt(Application.ExeName,'.ini')) then begin
     Ini := TIniFile.Create( ChangeFileExt( Application.ExeName, '.ini'));
    try
     Ini.WriteString('Time','Begin',bgt);
     Ini.WriteString('Time','End',ent);
    finally
     Ini.Free;
    end;
   end;
   RegWriteStr(HKEY_CURRENT_USER,'SOFTWARE\Microsoft\Windows\CurrentVersion\Run','Backup_GDB',ParamStr(0));
   iniF1:=ExtractFilePath(ParamStr(0))+'TransFiles.ini';  //TransFiles.ini
   d_BitBtn_Pusk.Tag:=10;
   d_BitBtn_Pusk.Enabled:=False;
   if FileExists(iniF1) then begin
   Ini:= TIniFile.Create(iniF1);
   try
    recConfigIni.sPathLogLocal:= SplitStr(Ini.ReadString('LocalDIR','LocalDIR',''));
    recConfigIni.sMask:=Ini.ReadString('Mask','Mask','');
    recConfigIni.sPathBackupServer:= SplitStr(Ini.ReadString('BackupDIR','BackupDIR',''));
    recConfigIni.sFirebird:=Ini.ReadString('Path','Firebird','');
    recConfigIni.sTimeBegin:=Ini.ReadString('Time','Begin','');
    recConfigIni.sTimeEnd:=Ini.ReadString('Time','End','');
    fireb:=SplitStr(recConfigIni.sFirebird);
    R:=TStringList.Create;
    ExtractStrings(['\'],['\'],PChar(recConfigIni.sPathLogLocal),R);
    for i:=0 to R.Count-1 do begin
        s1:=R.Strings[i];
        if i = 1 then begin
           y:=Pos(':',s);
        if y = 2 then
           s:=s+s1
        else s:=s+s1+':\'
        end
        else if i = 0 then begin
           y:=Pos(':',s1);
           if y = 2 then
           s:=s1+'\'
           else s:=s1+':';
           ips:=s1;
        end else begin
           y:=Pos(':',s);
           if y = 2 then
              s:=s+'\'+s1+'\'
           else s:=s+s1+'\';
        end;
        s:=AnsiReplaceStr(s, ' ', '');
    end;
    pth:=AnsiReplaceStr(s, ' ', '');
    R.Free;
    mMask:=recConfigIni.sMask+Copy(DateToStr(Now),7,4)+Copy(DateToStr(Now),4,2)+'.GDB'; //REG201603.GDB
   finally
      Ini.Free;
   end;
   end;
   FindFiles(recConfigIni.sPathBackupServer,'*',TST,True);
   if TST.Count > 0 then
   for i:=0 to TST.Count-1 do begin
       s:=TST.Strings[i];
       s:=SDStr(s);
       RemoveEmptyDirectories(s);
       lst1.Items.Add(StrTime+' -> '+s);
       if MyDirectoryDelete(s) then
          lst1.Items.Add(StrTime+ ' Папка '+s+' успешно удалена.')
       else lst1.Items.Add(StrTime+ ' Ошибка: папка '+s+' не удалена.');
   end;
   if TST.Count > 0 then lst1.Items.SaveToFile(ExtractFilePath(ParamStr(0))+'DirDel.log');
   d_Memo_Info.Lines.Add('Старт программы '+ DateToStr(Date)+' в '+ TimeToStr(GetTime));
   d_Memo_Info.Lines.Add(StrTime+Caption);
   d_Memo_Info.Lines.Add(StrTime+ 'Проверка связи с серверами ...');
   d_Memo_Info.Lines.Add(StrTime+ '==============================');
   iCount:= 0;
   bday:=DateToStr(Date-1);
   bcday:=DateToStr(Date-2);
   if not DirectoryExists(recConfigIni.sPathLogLocal) then ForceDirectories(recConfigIni.sPathLogLocal);
   if not DirectoryExists(recConfigIni.sPathBackupServer) then ForceDirectories(recConfigIni.sPathBackupServer);
   if LedSwich
   then begin
        bFlConnect:= True;
          if not DirectoryExists(recConfigIni.sPathLogLocal) then begin
             raise Exception.Create('Невозможно создать каталог '+ recConfigIni.sPathLogLocal);
             d_Memo_Info.Lines.Add(StrTime+ recConfigIni.sPathLogLocal);
             bFlConnect:= False;
          end else begin
               d_BitBtn_Pusk.Tag:=0;
               d_BitBtn_Pusk.Enabled:=True;
               d_Memo_Info.Lines.Add(StrTime+ 'Связь с сервером АРМВЗ - ОК');
          end;
   end else begin
        d_BitBtn_Pusk.Tag:=10;
        d_BitBtn_Pusk.Enabled:=False;
        bFlConnect:= False;
        d_Memo_Info.Lines.Add(StrTime+ 'Нет связи с сервером АРМВЗ: '+recConfigIni.sPathLogLocal);
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
               d_Memo_Info.Lines.Add(StrTime+ 'Связь с сервером Backup - ОК');
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
        d_Memo_Info.Lines.Add(StrTime+ 'Нет связи с сервером ЦПЗ: '+recConfigIni.sPathBackupServer);
   end;
   d_Memo_Info.Lines.Add(StrTime+ '==============================');
   TST.Free;
end;{Td_Form.OnCreate}

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
     if FindFirst(SplitStr(recConfigIni.sPathLogLocal)+ sMask, faAnyFile, SR)<>0
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

function FileExec( const CmdLine: String; bHide, bWait: Boolean): Boolean;
var
  StartupInfo : TStartupInfo;
  ProcessInfo : TProcessInformation;
begin
  FillChar(StartupInfo, SizeOf(TStartupInfo), 0);
  with StartupInfo do
  begin
    cb := SizeOf(TStartupInfo);
    dwFlags := STARTF_USESHOWWINDOW or STARTF_FORCEONFEEDBACK;
    if bHide then
       wShowWindow := SW_HIDE
    else wShowWindow := SW_SHOWNORMAL;
  end;
  Result := CreateProcess(nil, PChar(CmdLine), nil, nil, False,
            NORMAL_PRIORITY_CLASS, nil, nil, StartupInfo, ProcessInfo);
  if Result then
     CloseHandle(ProcessInfo.hThread);
  if bWait then
     if Result then
     begin
       WaitForInputIdle(ProcessInfo.hProcess, INFINITE);
       WaitForSingleObject(ProcessInfo.hProcess, INFINITE);
     end;
  if Result then
     CloseHandle(ProcessInfo.hProcess);
end;

procedure Td_Form.d_BitBtn_PuskClick(Sender: TObject);
label vx,vx1;
var
   SR: TSearchRec;
   sFName: TFileName;
   i: Integer;
   Retour: LongBool;
   StrL,Bat: TStringList;
   IBDB:TIBDataBase;
   IBInf:TIBDatabaseInfo;
   ConnectedUsers:TStringList;
   err: Integer;
begin
   tmr1.Enabled:=False;
   tmr2.Enabled:=False;
   if d_BitBtn_Pusk.Tag = 10 then Exit;
   d_BitBtn_Pusk.Enabled:=False;
   Label1.Caption:= '0';
   bFlConnect:= True;
   StrL := TStringList.Create;
   Bat := TStringList.Create;
   StrL.Delimiter := ',';
   StrL.DelimitedText := mMask;
   KillTask('gbak.exe');
   KillTask('gbak.exe');
   KillTask('gbak.exe');
   KillTask('gfix.exe');
   KillTask('gfix.exe');
   KillTask('gfix.exe');
   vx:
   for i := 0 to StrL.Count-1 do begin
   mMask:=StrL.Strings[i];
   with recConfigIni do
   try
   if FindFirst(SplitStr(recConfigIni.sPathLogLocal)+ mMask, faAnyFile, SR)=0 then
      sFName:=SR.Name;
      tmr1.Enabled:=False;
     if LedSwich and LedSwich2 then
     if FileExists(recConfigIni.sPathLogLocal+sFName) then begin
        d_Memo_Info.Lines.Add(StrTime+'Копирование файла ... '+sFName);
             repeat
                Application.ProcessMessages;
                if FileExists(recConfigIni.sPathLogLocal+sFName) then begin
                   Retour := False;
                   if DirectoryExists(recConfigIni.sPathLogLocal) then
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
                      d_Memo_Info.Lines.Add(StrTime + recConfigIni.sPathLogLocal+sFName + ' - Ok');
                      Application.ProcessMessages;
                    end else begin
                      bFlConnect:= False;
                      tmr2.Enabled:=True;
                      d_Memo_Info.Lines.Add(StrTime + 'Нет связи с сервером!! ');
                      Exit;
                    end;
                end;
             until ((FindNext(SR)<>0) or not bFlConnect);
             if not bFlConnect then d_Memo_Info.Lines.Add(StrTime + 'Остановлено пользователем');
             FindClose(sr);
             d_Memo_Info.Lines.Add(StrTime + 'Копирование файлов завершено');
             d_Memo_Info.Lines.Add(StrTime + 'Всего скопировано '+ IntToStr(iCount)+' файла');
             d_Memo_Info.Lines.Add('========================================');
             Application.ProcessMessages;
             if errs <> 0 then
             if FileExists(fireb+'gfix.exe') then begin
                tmr1.Enabled:=False;
                tmr2.Enabled:=False;
                errs:=0;
                if FileExists('c:\base_fix\IN.GDB') then
                   DeleteFile('c:\base_fix\IN.GDB');
                if FileExists('c:\base_fix\FIX.bat') then
                   DeleteFile('c:\base_fix\FIX.bat');
                if not FileExists('c:\base_fix\IN.GDB') then
                if not ForceDirectories('c:\base_fix') then
                   d_Memo_Info.Lines.Add(StrTime + 'Немогу создать каталог c:\base_fix')
                else
                if not CopyFileEx(
                       PChar(recConfigIni.sPathBackupServer+sFName),
                       PChar('c:\base_fix\IN.GDB'),
                       @CopyCallBack,
                       ProgressBar1,
                       @Retour,
                       COPY_FILE_RESTARTABLE)
                then d_Memo_Info.Lines.Add(StrTime + IntToStr(GetLastError));
                     d_Memo_Info.Lines.Add(StrTime+ 'Start fix c:\base_fix ...');
                   if FileExists('c:\base_fix\IN.GDB') then begin
                      Bat.Clear;
                      Bat.Add('SET ISC_USER=SYSDBA');
                      Bat.Add('SET ISC_PASSWORD=masterkey');
                      Bat.Add('"'+fireb+'gfix'+'"'+' -mend -full -ignore 127.0.0.1:c:\base_fix\IN.GDB');
                      Bat.Add('PAUSE');
                      Bat.Add('"'+fireb+'gbak'+'"'+' -b -v -ig -g 127.0.0.1:C:\base_fix\IN.GDB C:\base_fix\TMP.gbk');
                      Bat.Add('"'+fireb+'gbak'+'"'+' -c -v -p 4096 C:\base_fix\TMP.gbk 127.0.0.1:C:\base_fix\FIXED.gdb');
                      Bat.Add('DEL /Q TMP.gbk');
                      Bat.Add('PAUSE');
                      Bat.SaveToFile('c:\base_fix\FIX.bat');
                      Sleep(2000);
                      FileExec('c:\base_fix\FIX.bat',False,True);
                   end;
             end;
             if FileExists('C:\base_fix\FIXED.gdb') then begin
                tmr1.Enabled:=False;
                tmr2.Enabled:=False;
             if not FileExists(recConfigIni.sPathBackupServer+sFName+ '_' + DateToStr(now) + '.gdb') then
                CopyFile('C:\base_fix\FIXED.gdb',recConfigIni.sPathBackupServer+sFName+ '_' + DateToStr(now) + '.gdb');
                if FileExists(recConfigIni.sPathBackupServer+sFName+ '_' + DateToStr(now) + '.gdb') then begin
                   d_Memo_Info.Lines.Add(StrTime + 'Фикс базы успешно скопирован -> '+recConfigIni.sPathBackupServer+sFName+ '_' + DateToStr(now) + '.gdb');
                   DeleteFile('C:\base_fix\FIXED.gdb');
                end else begin
                   d_Memo_Info.Lines.Add(StrTime + 'Ошибка копирования базы -> '+recConfigIni.sPathBackupServer+sFName+ '_' + DateToStr(now) + '.gdb');
                   tmr2.Enabled:=True;
                end;
                goto vx1;
             end else
             if FileExists(recConfigIni.sPathBackupServer+sFName) then begin
                tmr1.Enabled:=False;
                tmr2.Enabled:=False;
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
                  fibBackUp.DatabaseName := '127.0.0.1:'+recConfigIni.sPathBackupServer+sFName; //pth+sFName;
                  fibBackUp.ServerName := 'localhost';
                  fibBackUp.BackupFile.Add(recConfigIni.sPathBackupServer + ExtractFileName(recConfigIni.sPathBackupServer+sFName)+ '_' + DateToStr(now) + '.gbk');
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
                       err:=AnsiPos('ERROR:',s);
                       if err > 0 then begin
                          tmr1.Enabled:=True;
                          tmr2.Enabled:=True;
                          d_Memo_Info.Lines.Add(StrTime + s);
                          Application.ProcessMessages;
                       end;
                    end;
                    if err > 0 then begin
                       tmr1.Enabled:=True;
                       tmr2.Enabled:=True;
                       d_Memo_Info.Lines.Add(StrTime + s);
                       Application.ProcessMessages;
                    end;
                    d_Memo_Info.Lines.Add(StrTime + '*************** Резервное копирование закончено ***************');
                    d_Memo_Info.Lines.SaveToFile(ExtractFilePath(ParamStr(0))+'Log.txt');
                    fibBackUp.Active := false;
                    Screen.Cursor := crDefault;
                    tmr2.Enabled:=True;
                  except
                    MessageDlg('Ошибка! Резервного копирования базы '+ sFName,mtError,[mbOk],0);
                    tmr2.Enabled:=True;
                  end;
               end;
              end else begin
                  ExitCode:=0;
                  d_Memo_Info.Lines.Add(StrTime + 'NO ACTIVE CONNECTIONS');
                  tmr1.Enabled:=True;
                  tmr2.Enabled:=True;
              end;
              except
              on E:Exception do
                begin
                  tmr1.Enabled:=True;
                  tmr2.Enabled:=True;
                  d_Memo_Info.Lines.Add(StrTime + 'Error='+E.Message);
                  ExitCode:=1;
                end;
              end;
             end;
             vx1:
             TrayIcon1.ShowBalloonHint('Внимание', 'Копирование файлов завершено!', bitInfo, 11);
             iCount:= 0;
             sFName:='';
             KillTask('gbak.exe');
             KillTask('gbak.exe');
             KillTask('gbak.exe');
             KillTask('gfix.exe');
             KillTask('gfix.exe');
             KillTask('gfix.exe');
     end else begin
        d_Memo_Info.Lines.Add(StrTime+'Нет файлов '+mMask+' для копирования!');
        tmr1.Enabled:=True;
        tmr2.Enabled:=True;
     end;
   except
      d_Memo_Info.Lines.Add(StrTime +'Ошибка копирования!');
      tmr1.Enabled:=True;
      tmr2.Enabled:=True;
   end;
   end;
   StrL.Free;
   Bat.Free;
   d_BitBtn_Pusk.Enabled:=True;
   d_Memo_Info.Lines.SaveToFile(ExtractFilePath(ParamStr(0))+'Log.txt');
   if err = 0 then begin
      tmr1.Enabled:=True;
      tmr2.Enabled:=True;
   end;
   if err > 0 then begin
      err:=0;
      errs:=1;
      tmr1.Enabled:=True;
      tmr2.Enabled:=True;
      MessageBox(0,PChar('Ошибка в базе данных!!!'+#13#10+s),PChar('Внимание'),13);
      d_Memo_Info.Lines.Add(StrTime +'Ошибка в базе данных -> '+#13#10+s);
      d_Memo_Info.Lines.Add(StrTime +'Ошибка! Резервного копирования базы!');
      d_Memo_Info.Lines.SaveToFile(ExtractFilePath(ParamStr(0))+'Log_ERROR.txt');
      goto vx;
   end;
   tmr2.Enabled:=True;
end;{Td_Form.d_BitBtn_PuskClick}

procedure Td_Form.FormDestroy(Sender: TObject);
begin
  KillTask('gbak.exe');
  KillTask('gbak.exe');
  KillTask('gbak.exe');
  KillTask('gfix.exe');
  KillTask('gfix.exe');
  KillTask('gfix.exe');
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
  i,y: Integer;
  vm: _SYSTEMTIME;
  vHour, vMin, vSec, vMm: Word;
  Ini: TIniFile;
begin
  DecodeTime(Now, vHour, vMin, vSec, vMm );
  vm.wHour := vHour ;
  vm.wMinute := vMin;
  vm.wSecond := vSec;
  vm.wMilliseconds := vMm;
  tm:=IntToStr(vm.wMinute+2);
  if recConfigIni.sTimeBegin = IntToStr(vm.wHour)+':'+IntToStr(vm.wMinute) then begin
     i:=Length(recConfigIni.sTimeBegin);
     y:=Pos(':',recConfigIni.sTimeBegin);
     s:=Copy(recConfigIni.sTimeBegin,1,y-1);
     s1:=Copy(recConfigIni.sTimeBegin,y+1,i);
     recConfigIni.sTimeBegin:=IntToStr(vm.wHour+5)+':'+IntToStr(vm.wMinute+1);
     recConfigIni.sTimeEnd:=IntToStr(vm.wHour+5)+':'+IntToStr(vm.wMinute+2);
  if not stime then begin
  if FileExists(ChangeFileExt(Application.ExeName,'.ini')) then begin
     Ini := TIniFile.Create( ChangeFileExt( Application.ExeName, '.ini'));
   try
     Ini.WriteString('Time','Begin',recConfigIni.sTimeBegin);
     Ini.WriteString('Time','End',recConfigIni.sTimeEnd);
   finally
     Ini.Free;
   end;
  end;
     Label1.Caption:='['+IntToStr(vm.wMinute)+']';
     tm:=IntToStr(vm.wMinute+2);
     tmr1.Enabled:=False;
     d_BitBtn_Pusk.Click;
  end;
     stime:=True;
  end;
  if recConfigIni.sTimeEnd = IntToStr(vm.wHour)+':'+IntToStr(vm.wMinute) then begin
     i:=Length(recConfigIni.sTimeEnd);
     y:=Pos(':',recConfigIni.sTimeEnd);
     s:=Copy(recConfigIni.sTimeEnd,1,y-1);
     s1:=Copy(recConfigIni.sTimeEnd,y+1,i);
     recConfigIni.sTimeBegin:=IntToStr(vm.wHour+5)+':'+IntToStr(vm.wMinute+1);
     recConfigIni.sTimeEnd:=IntToStr(vm.wHour+5)+':'+IntToStr(vm.wMinute+2);
  if not stime then begin
  if FileExists(ChangeFileExt(Application.ExeName,'.ini')) then begin
     Ini := TIniFile.Create( ChangeFileExt( Application.ExeName, '.ini'));
   try
     Ini.WriteString('Time','Begin',recConfigIni.sTimeBegin);
     Ini.WriteString('Time','End',recConfigIni.sTimeEnd);
   finally
     Ini.Free;
   end;
  end;
     Label1.Caption:='['+IntToStr(vm.wMinute)+']';
     tm:=IntToStr(vm.wMinute+2);
     tmr1.Enabled:=False;
     d_BitBtn_Pusk.Click;
  end;
     stime:=True;
  end;
   Label1.Caption:='['+IntToStr(vm.wSecond)+']';
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
          if not DirectoryExists(recConfigIni.sPathLogLocal) then begin
             if not ForceDirectories(recConfigIni.sPathLogLocal) then
             begin
               raise Exception.Create('Невозможно создать каталог '+ recConfigIni.sPathLogLocal);
               d_Memo_Info.Lines.Add(StrTime+ recConfigIni.sPathLogLocal);
               bFlConnect:= False;
               tmr2.Enabled:=True;
             end;
          end else begin
               d_BitBtn_Pusk.Tag:=0;
               d_BitBtn_Pusk.Enabled:=True;
               d_Memo_Info.Lines.Add(StrTime+ 'Связь с сервером АРМВЗ - ОК');
               tmr2.Enabled:=True;
          end;
   end
   else begin
        d_BitBtn_Pusk.Tag:=10;
        d_BitBtn_Pusk.Enabled:=False;
        bFlConnect:= False;
        d_Memo_Info.Lines.Add(StrTime+ 'Нет связи с сервером АРМВЗ: '+recConfigIni.sPathLogLocal);
        tmr2.Enabled:=True;
   end;
   if DirectoryExists(fireb) then
   if FileExists(fireb+'gfix.exe') then begin
      d_Memo_Info.Lines.Add(StrTime+ 'Папка Firebird: '+fireb+' - ОК');
      tmr2.Enabled:=True;
   end else begin
      d_Memo_Info.Lines.Add(StrTime+ 'Папка Firebird: '+fireb+' - ERROR');
      tmr2.Enabled:=True;
   end;
   if LedSwich2
   then begin
        bFlConnect:= True;
          if not DirectoryExists(recConfigIni.sPathBackupServer) then begin
             if not ForceDirectories(recConfigIni.sPathBackupServer) then
             begin
               raise Exception.Create('Невозможно создать каталог '+ recConfigIni.sPathBackupServer);
               d_Memo_Info.Lines.Add(StrTime+ recConfigIni.sPathBackupServer);
               tmr2.Enabled:=True;
               bFlConnect:= False;
             end;
          end else begin
               d_BitBtn_Pusk.Tag:=0;
               d_BitBtn_Pusk.Enabled:=True;
               d_Memo_Info.Lines.Add(StrTime+ 'Связь с сервером Backup - ОК');
               tmr2.Enabled:=True;
          end;
   end
   else begin
        d_BitBtn_Pusk.Tag:=10;
        d_BitBtn_Pusk.Enabled:=False;
        bFlConnect:= False;
        d_Memo_Info.Lines.Add(StrTime+ 'Нет связи с сервером Backup: '+recConfigIni.sPathBackupServer);
        tmr2.Enabled:=True;
   end;
   d_Memo_Info.Lines.Add(StrTime+ '==============================');
   tmr1.Enabled:=True;
   tmr2.Enabled:=True;
end;

procedure Td_Form.tmr2Timer(Sender: TObject);
var
  s: string;
  y: Integer;
  vm: _SYSTEMTIME;
  vHour, vMin, vSec, vMm: Word;
begin
  DecodeTime(Now, vHour, vMin, vSec, vMm );
  vm.wSecond := vSec;
  s:=IntToStr(vm.wSecond);
  y:=Pos('0',s);
  if y > 0 then begin
     tmr2.Enabled:=False;
     tmr1.Enabled:=True;
  end;
end;

procedure Td_Form.Exit1Click(Sender: TObject);
var
   Ini: TIniFile;
begin
  if FileExists(ChangeFileExt(Application.ExeName,'.ini')) then begin
    Ini := TIniFile.Create( ChangeFileExt( Application.ExeName, '.ini'));
   try
    Ini.WriteString('Path','Firebird',recConfigIni.sFirebird);
    Ini.WriteString('Time','Begin',recConfigIni.sTimeBegin);
    Ini.WriteString('Time','End',recConfigIni.sTimeEnd);
   finally
    Ini.Free;
   end;
  end;
  Application.Terminate;
end;

procedure Td_Form.FormActivate(Sender: TObject);
begin
  CreateFormInRightBottomCorner;
end;

end.
