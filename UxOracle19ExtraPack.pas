//----------------------------------------------------------------------------------------------------------------------
//
// ѕоддержка (частична€) Oracle 19 и выше дл€ DOA Delphi 5.0
// “акие модули до сих пор актуальны дл€ Nexign...
//
// “ипичное применение:
//   // ѕодставить выбранные параметры (здесь :prm_Yes и :prm_No)
//   ApplySqlStringVariables(myQuery, ['prm_Yes', 'prm_No']);
//   myQuery.Open();
//
//   // ѕодставить все строковые параметры
//   ApplySqlStringVariables(myQuery);
//   myQuery.Open();
//
// ѕомните, что ApplySqlStringVariables убирает переменную(ые) прив€зки, поэтому
//   1. ”худшаетс€ запрос
//   2. ћожет потребоватьс далнейша€ правка кода (провер€ть существет ли переменна€ и т.п)
//   3. ѕо факту, нужно убирать те переменные прив€зки, кторые в select: select ..., :prm_Name, ... from
//
//----------------------------------------------------------------------------------------------------------------------

unit
  UxOracle19ExtraPack;

interface

uses
  Windows, SysUtils, Classes, Db,

  Oracle,
  OracleData,

  UxStr,
  UxLexScan,
  UxOraBnd,
  UxUpts,
  UxUptsCmp,

  OracleXGlobalization;

// ¬ерси€ Oracle, например 12.1.0.2.0
function OracleVersion(ASession: TOracleSession): AnsiString; overload;

// ѕодставл€ет значени€ строковых параметров в текст запроса
procedure ApplySqlStringVariables(AQuery: TOracleDataSet); overload;

// ѕодставл€ет значени€ строковых параметров в текст запроса
procedure ApplySqlStringVariables(AQuery: TOracleDataSet; const VariablesToChange: array of string); overload;

// ѕодставл€ет значени€ строковых параметров в текст запроса
procedure ApplySqlStringVariables(AQuery: TOracleQuery); overload;

// ѕодставл€ет значени€ строковых параметров в текст запроса
procedure ApplySqlStringVariables(AQuery: TOracleQuery; const VariablesToChange: array of string); overload;

var
  // ¬ерси€ Oracle, начина€ с которой нужно начинать подстановку строковых параметров
  OracleVersionToStartApplyVariables: AnsiString = '19.0.0.0.0';

implementation

var
   KnownOracleVersion : AnsiString = '0.0.0.0.0';

function OracleVersion(ASession: TOracleSession): AnsiString;
var
  q: TOracleDataSet;
begin
  Result := KnownOracleVersion;

  if (Result <> '0.0.0.0.0') then
    Exit;

  q := TOracleDataSet.Create(nil);

  try
    q.Session := ASession;

    q.Sql.Text :=
      'select REGEXP_SUBSTR(Banner, ''[0-9]+(\.[0-9]+)+'') ' + #13#10 +
      '  from v$version                                    ' + #13#10 +
      ' where Banner like ''Oracle%''                      ';

    q.Open();

    if (q.Bof and q.Eof) then
      Result := '0.0.0.0.0'
    else
      Result := q.Fields[0].AsString;
  finally
    q.Free();
  end;

  KnownOracleVersion := Result;
end;

function SplitToVersion(value: AnsiString): TStringList;
var
  i: LongInt;
  St: AnsiString;
begin
  Result := TStringList.Create();

  STRxMakeLSByStr(Result, value, '.');

  for i := Result.Count - 1 downto 0 do
    begin
      St := AnsiUpperCase(Trim(Result[i]));
      St := Trim(STRxTrimLeft(St, '0'));

      if ((St = '') and (i >= Result.Count - 1)) then
        Result.Delete(i)
      else
        Result[i] := St;
    end;
end;

function CompareVersions(left: AnsiString; right: AnsiString): LongInt;
var
  LS1, LS2: TStringList;
  V1, V2: AnsiString;
  i, N: LongInt;
begin
  LS1 := SplitToVersion(left);
  LS2 := SplitToVersion(right);

  try
    if (LS1.Count < LS2.Count) then
      N := LS1.Count
    else
      N := LS2.Count;

    for i := 0 to N - 1 do
      begin
        V1 := LS1[i];
        V2 := LS2[i];

        if (Length(V1) > Length(V2)) then
          Result := 1
        else if (Length(V1) < Length(V2)) then
          Result := -1
        else
          Result := AnsiCompareText(V1, V2);

        if (Result <> 0) then
          Exit;
      end;

    if (LS1.Count < LS2.Count) then
      Result := -1
    else if (LS1.Count > LS2.Count) then
      Result := 1
    else
      Result := 0;
  finally
    LS2.Free();
    LS1.Free();
  end;
end;

// ѕодставл€ет значени€ строковых параметров в текст запроса
procedure ApplySqlStringVariables(AQuery: TOracleDataSet);
begin
  ApplySqlStringVariables(AQuery, []);
end;

// ѕодставл€ет значени€ строковых параметров в текст запроса
procedure ApplySqlStringVariables(AQuery: TOracleDataSet; const VariablesToChange: array of string); // TOracleQuery
var
  i, j, vt: LongInt;
  St: AnsiString;
  varName, varValue: AnsiString;
  AllVarNames: TStringList;
  AllPlaces: TBindPlace;
  ReplaceIndex: LongInt;

  VarNamesToReplace: TStringList;
  VarValuesToReplace: TStringList;
begin
  if (LocalizationIsDesignTime()) then
    Exit;

  if (not Assigned(AQuery)) then
    Exit;

  if (AQuery.VariableCount <= 0) then
    Exit;

  if (CompareVersions(OracleVersion(AQuery.Session), OracleVersionToStartApplyVariables) < 0) then
    Exit;

  VarNamesToReplace := TStringList.Create();
  VarValuesToReplace := TStringList.Create();

  try
    if (Length(VariablesToChange) <= 0) then
      for i := 0 to AQuery.VariableCount - 1 do
        begin
          vt := AQuery.VariableType(i);

          if ((vt = otString) or (vt = otDBChar) or (vt = otChar) or (vt = otPLSQLString)) then
            begin
              varValue := AQuery.GetVariable(i);

              if ((CoreLocalization.CurrentName = '') or (CoreLocalization.CurrentName = 'IV')) then
                varValue := EncodeToOracleChr(varValue)
              else
                varValue := EncodeToOracleUniStr(varValue);

              varName := Trim(AQuery.VariableName(i));

              if (Copy(varName, 1, 1) = ':') then
                varName := Trim(Copy(varName, 2, Length(varName)));

              VarValuesToReplace.Add(varValue);
              VarNamesToReplace.Add(varName);
            end;
          end
    else
      for i := 0 to Length(VariablesToChange) - 1 do
        begin
          varName := Trim(VariablesToChange[i]);

          if (Copy(varName, 1, 1) = ':') then
            varName := Trim(Copy(varName, 2, Length(varName)));

          j := AQuery.VariableIndex(varName);

          if (j < 0) then
            continue;

          vt := AQuery.VariableType(j);

          if (not ((vt = otString) or (vt = otDBChar) or (vt = otChar) or (vt = otPLSQLString))) then
            continue;

          varValue := AQuery.GetVariable(j);

          if ((CoreLocalization.CurrentName = '') or (CoreLocalization.CurrentName = 'IV')) then
            varValue := EncodeToOracleChr(varValue)
          else
            varValue := EncodeToOracleUniStr(varValue);

          VarValuesToReplace.Add(varValue);
          VarNamesToReplace.Add(varName);
        end;

    if (VarNamesToReplace.Count <= 0) then
      Exit;

    AllVarNames := TStringList.Create();

    try
      ParseForBindsEx(AQuery.SQL, AllVarNames, AllPlaces);

      // Zip().OrderByDescending() emulation

      for i := Length(AllPlaces) - 1 downto 0 do
        begin
          ReplaceIndex := -1;
          varName := AnsiUpperCase(AllVarNames[i]);

          for j := VarNamesToReplace.Count - 1 downto 0 do
            if (AnsiUpperCase(VarNamesToReplace[j]) = varName) then
              begin
                ReplaceIndex := j;

                break;
              end;

          if (ReplaceIndex >= 0) then
            begin
              St := AQuery.Sql[AllPlaces[i].StartPoint.Y];

              St := Copy(St, 1, AllPlaces[i].StartPoint.X) +
                    VarValuesToReplace[ReplaceIndex] +
                    Copy(St, AllPlaces[i].StopPoint.X + 1, Length(St));

              AQuery.Sql[AllPlaces[i].StartPoint.Y] := St;

              if (AQuery.VariableIndex(varName) >= 0) then
                AQuery.DeleteVariable(varName);
            end;
        end;

    finally
      AllVarNames.Free();
    end;

  finally
    VarValuesToReplace.Free();
    VarNamesToReplace.Free();
  end;
end;

// ѕодставл€ет значени€ строковых параметров в текст запроса
procedure ApplySqlStringVariables(AQuery: TOracleQuery);
begin
  ApplySqlStringVariables(AQuery, []);
end;

// ѕодставл€ет значени€ строковых параметров в текст запроса
procedure ApplySqlStringVariables(AQuery: TOracleQuery; const VariablesToChange: array of string); // TOracleQuery
var
  i, j, vt: LongInt;
  St: AnsiString;
  varName, varValue: AnsiString;
  AllVarNames: TStringList;
  AllPlaces: TBindPlace;
  ReplaceIndex: LongInt;

  VarNamesToReplace: TStringList;
  VarValuesToReplace: TStringList;
begin
  if (LocalizationIsDesignTime()) then
    Exit;

  if (not Assigned(AQuery)) then
    Exit;

  if (AQuery.VariableCount <= 0) then
    Exit;

  if (CompareVersions(OracleVersion(AQuery.Session), OracleVersionToStartApplyVariables) < 0) then
    Exit;

  VarNamesToReplace := TStringList.Create();
  VarValuesToReplace := TStringList.Create();

  try
    if (Length(VariablesToChange) <= 0) then
      for i := 0 to AQuery.VariableCount - 1 do
        begin
          vt := AQuery.VariableType(i);

          if ((vt = otString) or (vt = otDBChar) or (vt = otChar) or (vt = otPLSQLString)) then
            begin
              varValue := AQuery.GetVariable(i);

              if ((CoreLocalization.CurrentName = '') or (CoreLocalization.CurrentName = 'IV')) then
                varValue := EncodeToOracleChr(varValue)
              else
                varValue := EncodeToOracleUniStr(varValue);

              varName := Trim(AQuery.VariableName(i));

              if (Copy(varName, 1, 1) = ':') then
                varName := Trim(Copy(varName, 2, Length(varName)));

              VarValuesToReplace.Add(varValue);
              VarNamesToReplace.Add(varName);
            end;
          end
    else
      for i := 0 to Length(VariablesToChange) - 1 do
        begin
          varName := Trim(VariablesToChange[i]);

          if (Copy(varName, 1, 1) = ':') then
            varName := Trim(Copy(varName, 2, Length(varName)));

          j := AQuery.VariableIndex(varName);

          if (j < 0) then
            continue;

          vt := AQuery.VariableType(j);

          if (not ((vt = otString) or (vt = otDBChar) or (vt = otChar) or (vt = otPLSQLString))) then
            continue;

          varValue := AQuery.GetVariable(j);

          if ((CoreLocalization.CurrentName = '') or (CoreLocalization.CurrentName = 'IV')) then
            varValue := EncodeToOracleChr(varValue)
          else
            varValue := EncodeToOracleUniStr(varValue);

          VarValuesToReplace.Add(varValue);
          VarNamesToReplace.Add(varName);
        end;

    if (VarNamesToReplace.Count <= 0) then
      Exit;

    AllVarNames := TStringList.Create();

    try
      ParseForBindsEx(AQuery.SQL, AllVarNames, AllPlaces);

      // Zip().OrderByDescending() emulation

      for i := Length(AllPlaces) - 1 downto 0 do
        begin
          ReplaceIndex := -1;
          varName := AnsiUpperCase(AllVarNames[i]);

          for j := VarNamesToReplace.Count - 1 downto 0 do
            if (AnsiUpperCase(VarNamesToReplace[j]) = varName) then
              begin
                ReplaceIndex := j;

                break;
              end;

          if (ReplaceIndex >= 0) then
            begin
              St := AQuery.Sql[AllPlaces[i].StartPoint.Y];

              St := Copy(St, 1, AllPlaces[i].StartPoint.X) +
                    VarValuesToReplace[ReplaceIndex] +
                    Copy(St, AllPlaces[i].StopPoint.X + 1, Length(St));

              AQuery.Sql[AllPlaces[i].StartPoint.Y] := St;

              if (AQuery.VariableIndex(varName) >= 0) then
                AQuery.DeleteVariable(varName);
            end;
        end;

    finally
      AllVarNames.Free();
    end;

  finally
    VarValuesToReplace.Free();
    VarNamesToReplace.Free();
  end;
end;

end.
