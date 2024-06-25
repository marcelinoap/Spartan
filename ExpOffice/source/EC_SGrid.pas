unit EC_SGrid;

{$I CompConditionals.Inc}

interface

uses Classes, Grids,

     // TExportStringGrid Units
     EC_Main;

type
  TExportStringGrid = class(TSharewareExportComponent)
  private
    FStringGrid: TStringGrid;

    // *** Set Access Methods ***
    procedure SetStringGrid(Value: TStringGrid);
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;

    // *** Export Procedures ***
    function HasData: Boolean; override;

    function GetNumExportableItems: Integer; override;
    function CurrentItemSelected: Boolean; override;

    function CanExport(aIndex: Integer): Boolean;
    function GetColValue(aIndex: Integer): String; override;
    function GetExportColumns(var aColumns: TStringList): TStringList; override;
    procedure GetExportItemData(var aExportItem: TStringList); override;

  public
    constructor Create(AOwner: TComponent); override;  

  published
    property StringGrid: TStringGrid read FStringGrid write SetStringGrid;

    property ColWidths;
  end;

implementation

uses
  CommCtrl, Windows, SysUtils, Dialogs, Forms,

  // TExportStringGrid Units
  EC_Strings;

constructor TExportStringGrid.Create(AOwner: TComponent);
var
  i: Integer;
begin
  inherited;

  // Set Illegal Options
  FIllegalOptions := [ExportAllWhenNoneSelected];

  // Detect First StringGrid on Form (if there's not already a StringGrid assigned)
  if not Assigned(FStringGrid) then
  with TForm(Owner) do
    for i := 0 to ComponentCount - 1 do
      if Components[i] is TStringGrid then
      begin
        FStringGrid := TStringGrid(Components[i]);
        break;
      end;
end;

function TExportStringGrid.GetNumExportableItems: Integer;
begin
  // Calculate Number of Selected Items
  if (SelectedRowsOnly in FOptions) then
    with FStringGrid.Selection do
      FCount.SelectedItems := Bottom - Top + 1
  else FCount.SelectedItems := 0;

  // Return Number of Exportable Items
  if (SelectedRowsOnly in FOptions) and (FCount.SelectedItems > 0) then
    Result := FCount.SelectedItems
  else Result := FStringGrid.RowCount - 1;
end;

function TExportStringGrid.CurrentItemSelected: Boolean;
begin
  with FStringGrid.Selection do
    Result := (FCount.CurrentRow >= Top - 1) and (FCount.CurrentRow <= Bottom - 1);
end;

function TExportStringGrid.CanExport(aIndex: Integer): Boolean;
begin
  // Make Sure we don't export invisible cols (unless told to do so)
  if ExportInvisibleCols in FOptions then
    Result := True
  else Result := FStringGrid.ColWidths[aIndex] > 0;
end;

function TExportStringGrid.GetColValue(aIndex: Integer): String;
begin
  Result := FStringGrid.Cells[aIndex, 0];
end;

function TExportStringGrid.GetExportColumns(var aColumns: TStringList): TStringList;
var
  i: Integer;
begin
  Result := inherited GetExportColumns(aColumns);

  // Allocate StringList Capacity
  {$IFDEF Delphi3Up}
  aColumns.Capacity := FStringGrid.ColCount;
  {$ENDIF}

  for i := 0 to FStringGrid.ColCount - 1 do // For Each Column...
    if CanExport(i) then
      aColumns.Add(GetColumnCaption(i));  // ...Pass it on to the Descendant.
end;

procedure TExportStringGrid.GetExportItemData(var aExportItem: TStringList);
var
 i: Integer;
begin
  // Allocate StringList Capacity
  {$IFDEF Delphi3Up}
  aExportItem.Capacity := FExportColumns.Count;
  {$ENDIF}

  // All Fields in Item's Row for Columns that are Visible -> S
  with FStringGrid do
    for i := 0 to ColCount - 1 do
      if CanExport(i) then
        aExportItem.Add(Cells[i, FCount.CurrentRow + 1]);
end;

function TExportStringGrid.HasData: Boolean;
begin
  Result := True; // Assume Success

  // Abort under the following circumstances:
  if Assigned(FStringGrid) then
  with FStringGrid do
  begin
    if RowCount <= 1 then // If No Items to Export
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
  else raise Exception.CreateFmt(cNoComponentError, ['TStringGrid']);
end;

// *** Set Access Methods ***

procedure TExportStringGrid.SetStringGrid(Value: TStringGrid);
begin
  if Value <> FStringGrid then
  begin
    FStringGrid := Value;
    if Assigned(Value) then
      FStringGrid.FreeNotification(Self); // Add Notification if the StringGrid is Destroyed
  end;
end;

procedure TExportStringGrid.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited;

  // Assign nil to FStringGrid when the Assigned StringGrid is Freed
  if (AComponent = FStringGrid) and (Operation = opRemove) then
    FStringGrid := nil;
end;

end.
