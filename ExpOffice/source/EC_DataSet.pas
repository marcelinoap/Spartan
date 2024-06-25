unit EC_DataSet;

{$I CompConditionals.Inc}

interface

uses Classes, db, dbtables, variants,

     // TExportDataSet Units
     EC_Main;

const
  // Field Type Constants
  cDontUseDisplayFormatFieldTypes = [ftMemo, ftFmtMemo];

type
  TExportDataSet = class(TSharewareExportComponent)
  private
    function GetDateTimeFormat(aField: TField): TColType;

  protected
    FAllowMemos: Boolean;
    FDataSet: TDataSet;
    FBookmark: TBookmark;

    // *** Used only by Descendant Objects ***
    function GetRecordCount: Integer;

    // *** Abstract TExportDataSet Procedures ***
    function GetFieldCount: Integer; virtual; abstract;
    function GetField(i: Integer): TField; virtual; abstract;

    // *** Export Procedures ***
    procedure InitExport; override;
    procedure CleanUpExport; override;

    function AllowFieldType(aField: TField): Boolean;
    function CanExport(aField: TField): Boolean; virtual;
    function MoreExportItems: Boolean; override;

    function GetExportColumns(var aColumns: TStringList): TStringList; override;
    procedure GetExportItemData(var aExportItem: TStringList); override;

  public
    constructor Create(AOwner: TComponent); override;

  published
    property ColWidths;
  end;

implementation

uses
  CommCtrl, Windows, SysUtils, Dialogs, Forms,

  // TExportDataSet Units
  EC_Strings;

const
  // Field Type Constants
  cIllegalFieldTypes = [ftUnknown,
                        ftBytes,
                        ftVarBytes,
                        ftBLOB,
                        ftGraphic,
                        ftParadoxOLE,
                        ftdBASEOLE,
                        ftTypedBinary

                       {$IFDEF Delphi3Up}
                        ,ftCursor
                       {$ENDIF}

                       {$IFDEF Delphi4Up}
                        ,ftADT,
                        ftArray,
                        ftReference,
                        ftDataSet
                       {$ENDIF}
                        ];

constructor TExportDataSet.Create(AOwner: TComponent);
begin
  inherited;

  // Set Illegal Options
  FIllegalOptions := [ExportAllWhenNoneSelected];
end;

// *** Private Methods ***

function TExportDataSet.GetDateTimeFormat(aField: TField): TColType;
  procedure RemoveString(Substr: String; var S: String);
  var
    i: Integer;
  begin
    i := Pos(Substr, S);
    if i > 0 then
      Delete(S, i, Length(Substr));
  end;

  function HasMarker(const Markers: String; var S: String): Boolean;
  var
    i: Integer;
  begin
    Result := False; // Assume failure
    for i := 1 to Length(Markers) do
      if Pos(Markers[i], S) > 0 then
      begin
        Result := True;
        Break;
      end;
  end;

const
  cAmPm1 = 'am/pm';
  cAmPm2 = 'a/p';
  cAmPm3 = 'ampm';

  cDateMarkers = 'dmy';  // Dates contain these characters
  cTimeMarkers = 'hnst'; // Times contain these characters

var
  DF: String; // Display Format String
  HasDate, HasTime: Boolean;
begin
  Result := ctDateTime; // Assume DateTime format
  DF := Trim(LowerCase(TDateTimeField(aField).DisplayFormat)); // Get DisplayFormat String

  // Remove all AM/PM Symbols from Format String
  RemoveString(cAmPm1, DF);
  RemoveString(cAmPm2, DF);
  RemoveString(cAmPm3, DF);

  // Figure out TDateTimeFormat
  if (Length(DF) = 0) or (Pos('c', DF) > 0) then Exit; // If Blank or has 'c', return dtDateTime

  HasDate := HasMarker(cDateMarkers, DF); // See if it has a Date Value
  HasTime := HasMarker(cTimeMarkers, DF); // See if it has a Time Value

  if HasDate and HasTime then Exit;   // If has Date and Time Section, return ctDateTime
  if HasDate then Result := ctDateTime_ShowDateOnly else // else if has Date, return ctDateTime_ShowDateOnly
  if HasTime then Result := ctDateTime_ShowTimeOnly;     // else if has Time, return ctDateTime_ShowTimeOnly
end;

// *** Used only by Descendant Objects ***

function TExportDataSet.GetRecordCount: Integer;
// Return the value of the Record Count property, but only
// use it if it's a "verified" local table.
const
{$IFDEF Delphi4Up}
  cLocalTables = [ttDBase, ttParadox, ttFoxPro];
{$ELSE}
  cLocalTables = [ttDBase, ttParadox];
{$ENDIF}

var
  S: String;
begin
  Result := -1; // assume failure

  if FDataSet is TTable then
    with TTable(FDataSet) do
      if TableType in cLocalTables then
        Result := RecordCount
      else
        if TableType = ttDefault then // If Table Type not specified,
        begin
          S := Lowercase(ExtractFileExt(TableName)); // Get Extension of Table
          if (S = '.db') or (S = 'dbf') then         // if Local Table,
            Result := RecordCount;                   // use record count
        end;
end;

// *** Export Procedures ***

procedure TExportDataSet.InitExport;
begin
  inherited;

  FCount.SelectedItems := GetNumExportableItems;

  // Prepare DataSet
  with FDataSet do
  begin
    FBookmark := GetBookmark; // Save Current Position
    DisableControls;          // Disable Controls

    // Go to Top unless we're only exporting the current record
    if FCount.SelectedItems <> 1 then
      FDataSet.First;                    // Goto First Record
  end;
end;

procedure TExportDataSet.CleanUpExport;
begin
  inherited;

  // Restore DataSet
  with FDataSet do
  begin
    GotoBookmark(FBookmark); // Save Current Position
    EnableControls;          // Enable Controls
    FreeBookmark(FBookmark);
  end;
end;

function TExportDataSet.AllowFieldType(aField: TField): Boolean;
begin
  Result := not (aField.DataType in cIllegalFieldTypes); // If it's not an illegal field type,
end;

function TExportDataSet.CanExport(aField: TField): Boolean;
begin
  Result := AllowFieldType(aField);

  // Make Sure we don't export invisible cols (unless told to do so)
  if not (ExportInvisibleCols in FOptions) then
    Result := Result and aField.Visible;
end;

function TExportDataSet.MoreExportItems: Boolean;
begin
  if SelectedRowsOnly in FOptions then
    Result := inherited MoreExportItems
  else Result := not FDataSet.Eof;
end;

function TExportDataSet.GetExportColumns(var aColumns: TStringList): TStringList;
var
  ColumnInfo: TExportColumnInfo;
  aField: TField;
  i, j: Integer;
begin
  Result := inherited GetExportColumns(aColumns);

  // Allocate Columns Capacity
  {$IFDEF Delphi3Up}
  aColumns.Capacity := GetFieldCount;
  {$ENDIF}

  // Populate aColumns with the Export Columns
  for i := 0 to GetFieldCount - 1 do        // For Each Column...
  begin
    aField := GetField(i);

    if CanExport(aField) then         // ...that we're allowed to export...
    begin
      // Add to the columns list
      j := aColumns.Add(GetColumnCaption(i));

      // Attach Column Info data to aColumns
      ColumnInfo := TExportColumnInfo.Create;
      with ColumnInfo do
        case aField.DataType of
          ftBoolean : ColType := ctBoolean;

          ftString
          , ftFixedChar,
            ftWideString,
            ftLargeInt
          : ColType := ctString;

          ftSmallInt,
          ftInteger,
          ftWord,
          ftFloat,
          ftCurrency, // Leave these as ctNumber because the TField.Currency property
          ftBCD,      // will determine if they are displayed as currencies or not
          ftAutoInc : ColType := ctNumber;

          ftDate    : ColType := ctDate;
          ftTime    : ColType := ctTime;
          ftDateTime: ColType := GetDateTimeFormat(aField); // Set DateTimeFormat

          ftMemo,ftWideMemo : ColType := ctMemo;
          ftFmtMemo : ColType := ctFmtMemo;

        else
          ColType := ctUnknown;
        end;

      // If CurrencyField, Determine if ColType should be ctCurrency
      if (aField is TCurrencyField) and TCurrencyField(aField).Currency then
        ColumnInfo.ColType := ctCurrency;  // Display as Currency instead of Number.

      // If FloatField, Determine if ColType should be ctCurrency
      if (aField is TFloatField) and TFloatField(aField).Currency then
        ColumnInfo.ColType := ctCurrency;  // Display as Currency instead of Number.

      // If BCDField, Determine if ColType should be ctCurrency
      if (aField is TBCDField) and TBCDField(aField).Currency then
        ColumnInfo.ColType := ctCurrency;  // Display as Currency instead of Number.

      aColumns.Objects[j] := ColumnInfo;  // Add to aColumns StringList
    end;
  end;

  // Readjust Columns Capacity
  {$IFDEF Delphi3Up}
  with aColumns do Capacity := Count;
  {$ENDIF}
end;

procedure TExportDataSet.GetExportItemData(var aExportItem: TStringList);
var
 ItemInfo: TExportItemInfo;
 aField: TField;
 i, j: Integer;
begin
  // Allocate StringList Capacity
  {$IFDEF Delphi3Up}
  aExportItem.Capacity := GetFieldCount;
  {$ENDIF}

  // All Fields in Item's Row for Columns that are Visible -> S
  for i := 0 to GetFieldCount - 1 do       // For Each Column...
  begin
    aField := GetField(i); // Get Field to try to Export

    if CanExport(aField) then            // ...that we're allowed to export...
    begin
      // Add the appropriate item data
      j := aExportItem.Add(aField.AsString); // "Raw" Data

      // Add Display String Representation
      ItemInfo := TExportItemInfo.Create;
      with ItemInfo do
      begin
        // Determine if We Should Attach a Display String
        UseDisplayString := not (aField.DataType in cDontUseDisplayFormatFieldTypes);

        // Attach Display String (if using)
        if UseDisplayString then
          DisplayString := aField.DisplayText; // Set Display String

        // If Boolean Field, Assign Boolean Value
        if aField is TBooleanField then
          if aField.Value = null  then BooleanValue := bvNull else
          if aField.Value = True  then BooleanValue := bvTrue else
          if aField.Value = False then BooleanValue := bvFalse;
      end;
      aExportItem.Objects[j] := ItemInfo;  // Add to aExportItem StringList
    end;
  end;

  // Goto Next Record
  FDataSet.Next;
end;

end.
