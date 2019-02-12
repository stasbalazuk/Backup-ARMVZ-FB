f(*****************************************************************************)
(*    Назначение:                                                            *)
(*          Модуль предназначен для конвертирования платежного документа     *)
(*         (текстового файла V*.ZZZ) формата банка "Ощадбанк" (далее ФБО)    *)
(*         в документ - файл типа "Z" (далее ФТZ).  Модуль включает общие    *)
(*         функции формирования имени ФТZ, создания ФТZ, конвертирования     *)
(*         ФБА -> ФТZ, конвертирование каталога с файлами ФБА, определение   *)
(*         истинности файла ФБА.                                             *)
(*                                                                           *)
(*    Проект: Отдел информационных технологий УДППЗ "УКРПОШТА".              *)
(*    Исполнитель: Мальцев А.Б.                                              *)
(*****************************************************************************)


unit u_Convert;

interface
uses
   SysUtils, Classes, QDialogs, Controls, Graphics, Forms,
   v_CPZ_Dat;

type
{Структура данных Оборотно-сальдовый ряд из файла ФБO}
   TDataOSRZ = record
      Schet   : string[14];  {}
      TimeDat1: string[8];   {}
      TimeDat2: string[8];   {}
      OstatokB: string[19];  {}
      TypeR   : string[1];                   {}
      CountDoc: string[6];   {}
      SumDebOb: string[19];  {}
      CountCrd: string[6];   {}
      SumCrdOb: string[19];  {}
      OstatokE: string[19];  {}
      TypeOst : string[1];                   {}
   end;

{Структура данных платежного ряда из файла ФБO}
   TDataPRZ = record
      TypePlat : string[1];
      MFOInp   : string[9];
      SchetInp : string[14];
      EDROPUInp: string[14];
      NameInp  : string[40];
      MFOResv  : string[9];
      SchetRes : string[14];
      EDROPURes: string[14];
      NameRes  : string[40];
      VidPlat  : string[2];
      OperNom  : string[10];
      SumPlat  : string[19];
      DataInp  : string[8];
      TimeInp  : string[4];
      DataRes  : string[8];
      TimeRes  : string[4];
      NazPlat  : string[255];
//      NazPlat  : array [1..255] of char;
   end;

const
   chRasdel: char = #250; {Символ - Разделитель полей ФТZ}
   chNoInfo: char = #63;
   scTempDir = 'TempDat';
var
   s_FileNm_Inp,
   s_FileNm_Out : TFileName;      {Имена исходного и сконвертированого файлов}
   s_DirNm_Inp,
   s_DirNm_Out  : TFileName;
   FInpT : TextFile;               {Исходный (ФБА) и ..}
   FOutT: TextFile;               {..сконвертированный (ФТZ) файлы}
   sExtOut: string[5];
   sTempDirOut: TFileName;


   arNazPlat: string;   {Массив назначения платежа ФБА (динамический)}
   recDat1   : TDataOSRZ;      {Экземпляр структуры данных Оборотно-сальдового ряда ФБА}
   recDat2, recDat3: TDataPRZ; {Экземпляр структуры данных платежного ряда ФБА}
   blRwrt: boolean;

   bCharCount: byte;
   arDin: string;
   bPrgrssItm: byte = 0; {Индикатор прогресса от 0 до 1000}
   iMaxBlockCount: integer = 0;


{Формирование имени и пустого сконвертированного файла}
   function ConvertFileNameCreate(sFlNm, sDirInp: TFileName; sExt: string): TFileName;
{}
   function ConvertFileCreate(sFlNm: TFileName): TFileName;
{}
   function ConvertFile: integer;
{}
   function MaxIndicatCount: integer;
{}
   function MaxFileCount: integer;
{}
   function ConvertDir: integer;

   function FileZExists(sFlNm: TFileName): boolean;

   function DateFiltr(sFlNm: string): boolean;

//   function CopyFile(sFileNameInp, sFileNameOut: TFileName): integer;


implementation

uses
   f_DemoCnvrtr;

function FileZExists(sFlNm: TFileName): boolean;
begin
   Result:= False;
   if FileExists(sFlNm) then
   begin
      if (LowerCase(ExtractFileName(sFlNm)[1])='v') and (LowerCase(ExtractFileExt(sFlNm))='.zzz')
      then Result:= True;
   end;
end;{FileZExists}

function DelBegEmptyChar(sInp: string): string;
var
   sNew: string;
begin
   sNew:= sInp;
   while ((sNew[1]= ' ') and (Length(sNew)>1)) do Delete(sNew,1,1);
   Result:= sNew;
end;{DelBegEmptyChar}

function DelEndEmptyChar(sInp: string): string;
var
   ch: Char;
   sNew: string;
begin
   sNew:= sInp;
   ch:= #32;
   while sNew[Length(sNew)]= ch do Delete(sNew,Length(sNew),1);
   Result:= sNew;
end;{DelEndEmptyChar}


function DelEmptyChar(sInp: string): string;
begin
   Result:= DelEndEmptyChar(DelBegEmptyChar(sInp));
end;{DelEmptyChar}


function SplitStr(sInp: string): string;
begin
   Result:= chRasdel + Copy(sInp,1,40) + chRasdel + Copy(sInp,41,60)+
            chRasdel + Copy(sInp,101,60);
end;

function SplitFileName(sFlNm: string; var sPath: string; var sFlNmShrt: string;
                       var sFlExt: string): string;
var
   i: byte;
   s, s1: string;
begin
   s:='';
   for i:=length(sFlNm) downto 1 do
   begin
      if (length(sFlNm)-i) > 3 then sFlExt:=s;
      if sFlNm[i]='\' then sFlNmShrt:= s;
      s1:= GetCurrentDir;
   end;
   Result:= sFlNmShrt;

end;{SplitFileName}

{Формирование имени сконвертированного файла}
function ConvertFileNameCreate(sFlNm, sDirInp: TFileName; sExt: string): TFileName;
var
   s, sF, sPt: string;
begin
   if FileExists(sFlNm) then
   begin
      s_FileNm_Inp:= sFlNm;
      S:= s_FileNm_Inp;
      sF:= Copy(s_FileNm_Inp, length(S)-9, 6);
      sPt:= sDirInp;
      s_FileNm_Out:= sDirInp +'\Z'+ sF + sExt;
      Result:= s_FileNm_Out;
   end
   else
   begin
      if sFlNm <> ''
      then MessageDlg('Файл '+ sFlNm +' не существует!', mtWarning, [mbYes], 0, mbYes);
      Result:= '';
   end;
end;{ConvertFileNameCreate}

function ConvertFileCreate(sFlNm: TFileName): TFileName;
var
  NewFile: TFileStream;
  Msg: string;
begin
  if sFlNm <> '' then
  begin
     if FileExists(sFlNm) then
     begin
        Msg := 'Файл ' + sFlNm + ' уже существует. Заменить существующий файл?';
        if MessageDlg(Msg, mtConfirmation, [mbYes, mbNo], 0, mbNo) = mrNo
        then begin
                Result:= '';
                Exit;
             end;
            if Not DeleteFile(sFlNm) then
            begin
               MessageDlg('Невозможно изменить файл ' + s_FileNm_Inp +
                            '. Файл открыт только на чтение или открыт другим приложением.', mtError , [mbYes], 0);
               Result:= '';
               Exit;
            end;
     end;
     try
        try
           NewFile := TFileStream.Create(sFlNm, fmCreate or fmShareDenyNone);
        except
          MessageDlg('Невозможно создать файл '+s_FileNm_Inp, mtError , [mbYes], 0);
        end;
     finally
        FreeAndNil(NewFile);
     end;
     Result:= sFlNm;
  end else
  begin
     Result:= '';
     MessageDlg('Не открыт конвертируемый файл!', mtWarning, [mbYes], 0, mbYes);
  end;
end;{ConvertFileCreate}

function ReadBlokPlatRyad: integer;
begin
   Result:=0;
   try
     ReadLn(FInpT, recDat2.TypePlat, recDat2.MFOInp,
     recDat2.SchetInp, recDat2.EDROPUInp,
     recDat2.NameInp, recDat2.MFOResv,
     recDat2.SchetRes, recDat2.EDROPURes,
     recDat2.NameRes, recDat2.VidPlat,
     recDat2.OperNom, recDat2.SumPlat,
     recDat2.DataInp, recDat2.TimeInp,
     recDat2.DataRes, recDat2.TimeRes,
     recDat2.NazPlat);
   except
      MessageDlg('Невозможно открыть файл '+ s_FileNm_Inp, mtWarning, [mbYes], 0, mbYes);
      Result:=1;
   end;
end;

function MaxIndicatCount: integer;
begin
   iMaxBlockCount:= 0;
   if FileExists(s_FileNm_Inp)then
   begin
      try
         AssignFile(FInpT, s_FileNm_Inp);
		d_Frm_Main.Label2.Font.Color:= clRed;
		d_Frm_Main.Label2.Refresh;
         Reset(FInpT{, 1});
         iMaxBlockCount:= 0;
         while not EOF(FInpT) do
         begin
            ReadLn(FInpT);
            Inc(iMaxBlockCount);
         end;
         CloseFile(FInpT);
      except
         MessageDlg('Невозможно открыть файл '+ s_FileNm_Inp, mtWarning, [mbYes], 0, mbYes);
      end;
      Result:= iMaxBlockCount;
   end
   else Result:= 0;

end;{MaxIndicatCount}
(*
function CopyFile(sFileNameInp, sFileNameOut: TFileName): integer;
var
  NewFileName: string;
  Msg: string;
  NewFile: TFileStream;
  OldFile: TFileStream;
begin
  Result:= 0;
  NewFileName := sFileNameOut;
{}  Msg := Format('Copy %s to %s?', [sFileNameInp, NewFileName]);
{}  if MessageDlg(Msg, mtCustom, mbOKCancel, 0) = mrOK then
  begin
    OldFile := TFileStream.Create(sFileNameInp, fmOpenRead or fmShareDenyWrite);
    try
      NewFile := TFileStream.Create(NewFileName, fmCreate or fmShareDenyRead);
      try
        NewFile.CopyFrom(OldFile, OldFile.Size);
      finally
        FreeAndNil(NewFile);
      end;
    finally
      FreeAndNil(OldFile);
    end;
   end;
end;{CopyFile}
*)
function MaxFileCount: integer;
var
   iCnt: integer;
   SR: TSearchRec;
begin
   iCnt:=0;
   if FindFirst(s_DirNm_Inp+'\*.*', 32, SR) = 0 then
   begin
     repeat
       if FileZExists(s_DirNm_Inp+'\'+ SR.Name) then
       begin
          if DateFiltr(SR.Name) then Inc(iCnt);
       end;
     until FindNext(SR) <> 0;
     FindClose(SR);
     Result:= iCnt;
   end else Result:=0;
end;{MaxFileCount}

function ConvertFile: integer;
var
   ch        : Char;
   sBuf1     : string;
   sDop: string;
   i: integer;
   sFNOut, sF: TFileName;

begin
   Result:= 0;
   bPrgrssItm:= 0;

   try
   AssignFile(FInpT, s_FileNm_Inp);
   Reset(FInpT);
   if iMaxBlockCount > 2 then d_Frm_Main.Gauge1.Progress:= 2;
   while not Eof(FInpT) do
   begin
       Read(FInpT, ch);
       case ch of
            '0': begin
                 with recDat1 do
                 ReadLn(FInpT,Schet,TimeDat1,TimeDat2,
                             OstatokB,TypeR,CountDoc,
                             SumDebOb,CountCrd,SumCrdOb,
                             OstatokE,TypeOst);

                    Inc(bPrgrssItm);
                    d_Frm_Main.Gauge1.Progress:= bPrgrssItm;
{+Bor}
                    if (StrToInt(String(recDat1.CountDoc)) <> 0) or (StrToInt(String(recDat1.CountCrd)) <> 0) then
                    begin

{Создание имени файла}
                       for i:=1 to btCntSchet do
                       begin
                       sDop:= Copy(arCPZ[i], 1, 12);
                         if DelEmptyChar(recDat1.Schet) = Copy(arCPZ[i], 1, 12) then
                         begin
                            sExtOut:= Copy(arCPZ[i], 13, 5);
                            Break;
                         end;
                       end;
                       if i>btCntSchet then sExtOut:= '1.122';
                       if FileExists(sFNOut)then CloseFile(FOutT);

                       sFNOut:= ConvertFileNameCreate(s_FileNm_Inp, sTempDirOut, sExtOut);

                       if not FileExists(sFNOut)
                       then sF:= ConvertFileCreate(sFNOut);
                       AssignFile(FOutT, sFNOut);
                       Append(FOutT);

                       recDat1.Schet:= Copy(recDat1.Schet, 1, 8);

                       sBuf1:= recDat1.Schet + chRasdel;
                       sBuf1:= sBuf1+chRasdel+chRasdel+chRasdel+chRasdel+chRasdel+chRasdel
                            +chRasdel+chRasdel+chRasdel+chRasdel+chRasdel+chRasdel;
                       Writeln(FOutT, sBuf1);
                    end;
                 end;
            '1': begin
                    ReadBlokPlatRyad;  {Считывание Платежного ряда}
                    Inc(bPrgrssItm);
                    d_Frm_Main.Gauge1.Progress:= bPrgrssItm;
                    {Запись в конверционный файл}
                    sBuf1:= chRasdel + chRasdel;
                    if recDat2.TypePlat = '+' then  ch:= '2' else
                    if recDat2.TypePlat = '-' then  ch:= '1' else;
                    sBuf1:= sBuf1 + ch + chRasdel;
                    if ch = '2' then
                    begin
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.MFOInp)+ chRasdel;
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.SchetInp)+ chRasdel;
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.OperNom)+ chRasdel;
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.DataInp)+ chRasdel;

                       sBuf1:= sBuf1 + '0' + chRasdel;
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.SumPlat)+ chRasdel;

                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.DataInp)+ chRasdel;
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.TimeInp)+ chRasdel;
                       sBuf1:= sBuf1 + DelEmptyChar(recDat2.NameInp)+ chRasdel;
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.EDROPUInp)+ chRasdel;
                       sBuf1:= sBuf1 + chRasdel + chRasdel + chRasdel + chRasdel + chRasdel;
                       sBuf1:= sBuf1 + SplitStr(DelBegEmptyChar(recDat2.NazPlat));

                    end else
                    if ch = '1' then
                    begin
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.MFOResv)+ chRasdel;
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.SchetRes)+ chRasdel;
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.OperNom)+ chRasdel;
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.DataRes)+ chRasdel;

                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.SumPlat)+ chRasdel;
                       sBuf1:= sBuf1 + '0' + chRasdel;

                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.DataRes)+ chRasdel;
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.TimeRes)+ chRasdel;
                       sBuf1:= sBuf1 + DelEmptyChar(recDat2.NameRes)+ chRasdel;
                       sBuf1:= sBuf1 + DelBegEmptyChar(recDat2.EDROPURes)+ chRasdel;
                       sBuf1:= sBuf1 + chRasdel + chRasdel + chRasdel + chRasdel + chRasdel;
                       sBuf1:= sBuf1 + SplitStr(DelBegEmptyChar(recDat2.NazPlat));
                    end;

                    SetLength(arNazPlat, 0);
                    Writeln(FOutT, sBuf1);
                 end;
            '2': begin
                    ReadBlokPlatRyad;
                    Inc(bPrgrssItm);
                    d_Frm_Main.Gauge1.Progress:= bPrgrssItm;
                    SetLength(arNazPlat, 0);
                 end;
       end;
   end;
   CloseFile(FInpT);
   CloseFile(FOutT);

   d_Frm_Main.Gauge1.Progress:= d_Frm_Main.Gauge1.MaxValue;
   except
      Result:= 1;
      MessageDlg('Не возможно создать файл '+s_FileNm_Out, mtError, [mbYes], 0);
   end;
end;{ConvertFile}

/////////////////////////////////////////////

function DateFiltr(sFlNm: string): boolean;
var
   sDate: string;
   sD, sM, sY: string[2];
   ADate: TDateTime;
begin
   Result:= False;
   sD:= Copy(sFlNm, 6, 2); sM:= Copy(sFlNm, 4, 2); sY:= Copy(sFlNm, 2, 2);
   sDate:= sD+'.'+sM+'.20'+sY;
   try
     ADate:= StrToDate(sDate);
   except
     MessageDlg('В названии файла  '+ sFlNm +'  не допустимая дата!', mtError, [mbYes], 0, mbYes);
     Exit;
   end;
   if (Trunc(d_Frm_Main.DateTimePicker1.Date) <= ADate) and (ADate <= {Trunc}(d_Frm_Main.DateTimePicker2.Date))
   then Result:= True;
end;{DateFiltr}

////////////////////////////////////////

function ConvertDir: integer;
var
   iCnt: integer;
   SR: TSearchRec;
   sFlNm, sFlNmNew: TFileName;
begin
   iCnt:=0;
   if FindFirst(s_DirNm_Inp+'\*.*', {faArchive}32, sr) = 0 then
   begin
     repeat
       sFlNm:= s_DirNm_Inp+'\'+SR.Name;
       if FileZExists(sFlNm) then
       begin
          if DateFiltr(SR.Name) then
          begin
             s_FileNm_Inp:= sFlNm;
//             s_FileNm_Out:= ConvertFileNameCreate(s_FileNm_Inp, s_DirNm_Out, d_Frm_Main.Edit1.Text,
//                                                                             d_Frm_Main.Edit2.Text);
             d_Frm_Main.Gauge1.MaxValue:= MaxIndicatCount;
             d_Frm_Main.Gauge1.Progress:= 0;
             if ConvertFile = 0 then Inc(iCnt);
          end;
          d_Frm_Main.Gauge2.Progress:= iCnt;
       end;
     until FindNext(SR) <> 0;
     FindClose(SR);
     Result:= iCnt;
   end else Result:=0;

{Копирование на сервер}
   if FindFirst(sTempDirOut+'\*.*', 32, sr) = 0 then
   begin
     repeat
       sFlNm:= sTempDirOut +'\'+ SR.Name;
       sFlNmNew:= s_DirNm_Out +'\'+ SR.Name;
       if FileExists(sFlNmNew) then
       if not DeleteFile(sFlNmNew) then
	       raise Exception.Create(' Невозможно заменить существующий файл '+ sFlNmNew +' ! ');
       if not RenameFile(sFlNm, sFlNmNew) then
    	   raise Exception.Create(' Невозможно скопировать файл '+ sFlNm+' в '+ s_DirNm_Out+' ! ');
     until FindNext(SR) <> 0;
     FindClose(SR);
   end else Result:=0;
end;{ConvertDir}

End.





