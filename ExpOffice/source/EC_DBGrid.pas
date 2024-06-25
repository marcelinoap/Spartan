unit EC_DBGrid;

{$I CompConditionals.Inc}

interface

uses Classes, db, dbgrids,

     // TExportDBGrid Units
     EC_DataSet;

type
  TExportDBGrid = class(TExportDataSet)
  private
    FDBGrid: TDBGrid;

    // *** Set Access Methods ***
    procedure SetDBGrid(Value: TDBGrid);
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;

    // *** Abstract TExportDataSet Procedures ***
    function GetFieldCount: Integer; override;
    function GetField(i: Integer): TField; override;

    // *** Export Procedures ***
    procedure InitExport; override;
    function CanExport(aField: TField): Boolean; override;

    function HasData: Boolean; override;

    function GetRealColIndex(var aColIndex: Integer): Integer;
    function GetDefaultColWidth(ColIndex: Integer): Integer; override;

    function GetNumExportableItems: Integer; override;

    function CurrentItemSelected: Boolean; override;
  {$IFDEF Delphi3Up}
    procedure SkipToNextItem; override;
  {$ENDIF}

    function GetColValue(aIndex: Integer): String; override;

  public
    constructor Create(AOwner: TComponent); override;

  published
    property DBGrid: TDBGrid read FDBGrid write SetDBGrid;
  end;

implementation

uses
  CommCtrl, Windows, SysUtils, Dialogs, Forms,

  // TExportDBGrid Units
  EC_Strings, EC_Main;

constructor TExportDBGrid.Create(AOwner: TComponent);
var
  i: Integer;
begin
  inherited;

  // Detect First DBGrid on Form (if there's not already a DBGrid assigned)
  if not Assigned(FDBGrid) then
  with TForm(Owner) do
    for i := 0 to ComponentCount - 1 do
      if Components[i] is TDBGrid then
      begin
        FDBGrid := TDBGrid(Components[i]);
        break;
      end;
end;

// *** Abstract TExportDataSet Procedures ***

function TExportDBGrid.GetFieldCount: Integer;
begin
  Result := FDBGrid.FieldCount;
end;

function TExportDBGrid.GetField(i: Integer): TField;
begin
  Result := FDBGrid.Fields[i];
end;

// *** Export Procedures ***

procedure TExportDBGrid.InitExport;
begin
  FDataSet := FDBGrid.DataSource.DataSet; // Set Data Set

  inherited;
end;

function TExportDBGrid.CanExport(aField: TField): Boolean;
{$IFDEF Delphi4Up}
var
  i: Integer;
{$ENDIF}
begin
  Result := AllowFieldType(aField);

  // Make Sure we don't export invisible cols (unless told to do so)
  if not (ExportInvisibleCols in FOptions) then
  begin
    // Export field only if both it's Visible property and it's column's visible property is true
    Result := Result and aField.Visible; // See if this field is visible
   {$IFDEF Delphi4Up}
    // Get Number of Columns
    with FDBGrid do
      for i := 0 to Columns.Count - 1 do
        if Columns[i].Field = aField then
          Result := Result and Columns[i].Visible;
   {$ENDIF}
  end;
end;

function TExportDBGrid.GetRealColIndex(var aColIndex: Integer): Integer;
var
  i, ExportableCols: Integer;
begin
  // Initialize
  ExportableCols := 0;
  Result         := -1; // Return -1 on failure

  // Count number of UnExportable Columns from the start until the column we're after
  with FDBGrid do
    for i := 0 to Columns.Count - 1 do
      if CanExport(Columns[i].Field) then
      begin
        Inc(ExportableCols);
        if ExportableCols > aColIndex then
        begin
          Result := i; // Return "Real" Col Index Value
          Break;
        end;
      end;
end;

function TExportDBGrid.GetDefaultColWidth(ColIndex: Integer): Integer;
var
  RealColIndex: Integer;
begin
  // Initialize
  Result := 0; // assume failure

  // Calculate Real ColIndex, they can be different values because some fields
  // may already be invisible or be of a non-exportable data type.
  RealColIndex := GetRealColIndex(ColIndex);
  if RealColIndex = -1 then Exit;

  with FDBGrid do
    if RealColIndex < Columns.Count then
      Result := Columns[RealColIndex].Field.DisplayWidth;
end;

function TExportDBGrid.GetNumExportableItems: Integer;
begin
  // Return Number of Exportable Items
  if SelectedRowsOnly in FOptions then // If SelectedRowsOnly Flagged,
  begin
  {$IFDEF Delphi3Up}
    if FDBGrid.SelectedRows.Count > 0 then // If multiple rows selected,
      Result := FDBGrid.SelectedRows.Count // find out how many,
    else {$ENDIF} Result := 1;             // if only 1 then, return 1.
  end
  else Result := GetRecordCount;
end;

function TExportDBGrid.CurrentItemSelected: Boolean;
begin
  Result := (FCount.SelectedItems = 1)
           {$IFDEF Delphi3Up} or FDBGrid.SelectedRows.CurrentRowSelected {$ENDIF};
end;

function TExportDBGrid.GetColValue(aIndex: Integer): String;
begin
  Result := FDBGrid.Columns[aIndex].Title.Caption;
end;

{$IFDEF Delphi3Up}
procedure TExportDBGrid.SkipToNextItem;
begin
  inherited;

  FDataSet.Next;
end;
{$ENDIF}

function TExportDBGrid.HasData: Boolean;
  function NoRecords: Boolean;
  begin
    with FDBGrid.DataSource.DataSet do
      Result := Bof and Eof; // If both Beginning and End of File, there're no records
  end;

begin
  Result := True; // Assume Success

  // DataSet must be assigned DBGrid.DataSource.DataSet

  // Abort under the following circumstances:
  if Assigned(FDBGrid) then
  begin
    if Assigned(FDBGrid.DataSource) then
    begin
      if Assigned(FDBGrid.DataSource.DataSet) then
      begin
        if not FDBGrid.DataSource.DataSet.Active then // If DataSet not Active,
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
        else raise Exception.CreateFmt(cNoDataSet, [FDBGrid.DataSource.Name]);
      end
      else raise Exception.CreateFmt(cNoDataSource, [FDBGrid.Name]);
  end
  else raise Exception.CreateFmt(cNoComponentError, ['TDBGrid']);
end;

// *** Set Access Methods ***

procedure TExportDBGrid.SetDBGrid(Value: TDBGrid);
begin
  if Value <> FDBGrid then
  begin
    FDBGrid := Value;
    if Assigned(Value) then
      FDBGrid.FreeNotification(Self); // Add Notification if the DBGrid is Destroyed
  end;
end;

procedure TExportDBGrid.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited;

  // Assign nil to FDBGrid when the Assigned DBGrid is Freed
  if (AComponent = FDBGrid) and (Operation = opRemove) then
    FDBGrid := nil;
end;

end.
