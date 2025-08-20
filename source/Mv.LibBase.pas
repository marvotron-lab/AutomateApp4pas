unit Mv.LibBase;

{
    Mv-Lib functions, that are used in this project
}


interface

uses
    System.SysUtils;

//Mv.BaseUtils
function Cat(const AStr1, AStr2, AConcatenator: string): string;
function ECat(const AStr1, AStr2: string; AConcatenator: string = ''): string;

//Mv.HexUtils
function HexToBytes(AValue: string): TBytes;

implementation


{ ConCATenation of 2 strings with AConcatenator if both are not empty (also @see Cat)
  e.g.:
    Cat('Hello', 'world', #13#10) -> 'Hello'#13#10'world' //#
    Cat('', 'world', #13#10) -> 'world' //#
------------------------------------------------------------------------------------------------------------------}
function Cat(const AStr1, AStr2, AConcatenator: string): string;
begin
    if AStr2 = '' then
      Result := AStr1
    else if AStr1 = '' then
      Result := AStr2
    else
      Result := AStr1 + AConcatenator + AStr2;
end;

{ ConCATenate 2 strings with AConcatenator, but return Empty string, if one is empty (also @see Cat)
  e.g.:
    ECat('Phone:', '1234', ' ') -> 'Phone: 1234' //#
    ECat('Phone: ', '1234') -> 'Phone: 1234' //#
    ECat('Phone:', '', ' ') -> '' //#
------------------------------------------------------------------------------------------------------------------}
function ECat(const AStr1, AStr2: string; AConcatenator: string = ''): string;
begin
    if (AStr1 = '') or (AStr2 = '') then
      Result := ''
    else
      Result := AStr1 + AConcatenator + AStr2;
end;

{ @param AValue Hex string, eg. '34DFD8C7'
  @return Byte array with the bytes represented by the respective values in the string,
    eg. [$34], [$DF], [$D8], [$C7]. Reverses @see BytesToHex
------------------------------------------------------------------------------------------------------------------}
function HexToBytes(AValue: string): TBytes;
var
    I: Integer;
    Hex: string;
    Val: Integer;
    Pos: Integer;
begin
    SetLength(Result, Round(Length(AValue) / 2));  //every two characters encode one byte
    I := 1;
    Pos := 0;
    while I < Length(AValue) do
    begin
        Hex := Copy(AValue, I, 2);
        Val := StrToInt('$' + Hex);
        Result[Pos] := Val;
        I := I + 2;
        Inc(Pos);
    end;
end;

end.

