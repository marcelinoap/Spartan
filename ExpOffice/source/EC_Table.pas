unit EC_Table;

{$I CompConditionals.Inc}

interface

uses Classes, db,
                                      
     // TExportTable Units
     EC_DataSet;

type
  TExportTable = class(TExportDataSet)
  private
    // *** Set Access Methods ***
    procedure SetDataSet(Value: TDataSet);
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;

    // *** Abstract TExportDataSet Procedures ***
    function GetFieldCount: Integer; override;
    function GetField(i: Integer): TField; override;

    // *** Export Procedures ***
    function HasData: Boolean; override;

    function GetRealColIndex(var aColIndex: Integer): Integer;
    function GetDefaultColWidth(ColIndex: Integer): Integer; override;

    function GetNumExportableItems: Integer; override;
    function CurrentItemSelected: Boolean; override;

    function GetColValue(aIndex: Integer): String; override;

  public
    constructor Create(AOwner: TComponent); override;

  published
    property DataSet: TDataSet read FDataSet write SetDataSet;
  end;

implementation

uses
  CommCtrl, Windows, SysUtils, Dialogs, Forms,

  // TExportTable Units
  EC_Strings, EC_Main;

constructor TExportTable.Create(AOwner: TComponent);
var
  i: Integer;
begin
  inherited;

  // Detect First DataSet on Form (if there's not already a DataSet assigned)
  if not Assigned(FDataSet) then
  with TForm(Owner) do
    for i := 0 to ComponentCount - 1 do
      if Components[i] is TDataSet then
      begin
        FDataSet := TDataSet(Components[i]);
        break;
      end;
end;

// *** Abstract TExportDataSet Procedures ***

function TExportTable.GetFieldCount: Integer;
begin
  Result := FDataSet.FieldCount;
end;

function TExportTable.GetField(i: Integer): TField;
begin
  Result := FDataSet.Fields[i];
end;

// *** Export Procedures ***

function TExportTable.GetRealColIndex(var aColIndex: Integer): Integer;
var
  i, ExportableCols: Integer;
begin
  // Initialize
  ExportableCols := 0;
  Result         := -1; // Return -1 on failure

  // Count number of UnExportable Columns from the start until the column we're after
  with FDataSet do
    for i := 0 to FieldCount - 1 do
      if CanExport(Fields[i]) then
      begin
        Inc(ExportableCols);
        if ExportableCols > aColIndex then
        begin
          Result := i; // Return "Real" Col Index Value
          Break;
        end;
      end;
end;

function TExportTable.GetDefaultColWidth(ColIndex: Integer): Integer;
var
  RealColIndex: Integer;
begin
  // Initialize
  Result := 0; // assume failure

  // Calculate Real ColIndex, they can be different values because some fields
  // may already be invisible or be of a non-exportable data type.
  RealColIndex := GetRealColIndex(ColIndex);
  if RealColIndex = -1 then Exit;

  with FDataSet do
    if RealColIndex < FieldCount then
      Result := Fields[RealColIndex].DisplayWidth;
end;

function TExportTable.GetNumExportableItems: Integer;
begin
  // Return Number of Exportable Items
  if SelectedRowsOnly in FOptions then // If SelectedRowsOnly Flagged,
    Result := 1                        // then, return 1 (for the current row)
  else Result := GetRecordCount;
end;

function TExportTable.CurrentItemSelected: Boolean;
begin                            
  Result := FCount.SelectedItems = 1;
end;

function TExportTable.GetColValue(aIndex: Integer): String;
begin
 Result := FDataSet.Fields[aIndex].DisplayName;
end;

function TExportTable.HasData: Boolean;
  function NoRecords: Boolean;
  begin
    with FDataSet do
      Result := Bof and Eof; // If both Beginning and End of File, there're no records
  end;

begin
  Result := True; // Assume Success

  // Abort under the following circumstances:
  if Assigned(FDataSet) then
  with FDataSet do
  begin
    if not Active then // If DataSet not Active,
      raise Exception.Create(cNotActiveError)
    else
      if NoRecords then // If No Items to Export
      begin
        MessageDlg(cNoDataError, mtWarning, [mbOk], 0);
        Result := False;
        Exit;
      end
      else
        if not HasExportableColumns then // If No Columns to Export
        begin
          Result := False;
          Exit;
        end;
  end
  else raise Exception.CreateFmt(cNoComponentError, ['TDataSet']);
end;

// *** Set Access Methods ***

procedure TExportTable.SetDataSet(Value: TDataSet);
begin
  if Value <> FDataSet then
  begin
    FDataSet := Value;
    if Assigned(Value) then
      FDataSet.FreeNotification(Self); // Add Notification if the DataSet is Destroyed
  end;
end;

procedure TExportTable.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited;

  // Assign nil to FDataSet when the Assigned DataSet is Freed
  if (AComponent = FDataSet) and (Operation = opRemove) then
    FDataSet := nil;
end;

end.
