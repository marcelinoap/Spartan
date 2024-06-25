unit EC_LView;

interface

uses Classes, ComCtrls,

     // TExportListView Units
     EC_Main;

const
  cMaxArraySize = High(Integer) div SizeOf(Integer) - 1;
type
  AInteger = array [0..cMaxArraySize] of Integer;

  TExportListView = class(TSharewareExportComponent)
  private
    FListView: TListView;

    ColOrderArray: ^AInteger;

    function ColAllowed(ListCol: TListColumn): Boolean; // Returns True if the Column should be exported
    function GetColIndex(aIndex: Integer): Integer;

    // *** Set Access Methods ***
    procedure SetListView(Value: TListView);
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;

    // *** Export Procedures ***
    procedure InitExport; override;
    procedure CleanUpExport; override;

    function HasData: Boolean; override;
    function GetColValue(aIndex: Integer): String; override;

    function GetNumExportableItems: Integer; override;
    function CurrentItemSelected: Boolean; override;

    function GetExportColumns(var aColumns: TStringList): TStringList; override;
    procedure GetExportItemData(var aExportItem: TStringList); override;

  public
    constructor Create(AOwner: TComponent); override;

  published
    property ListView: TListView read FListView write SetListView;

    property ColWidths;
  end;

implementation

uses
  CommCtrl, Windows, SysUtils, Dialogs, Forms,

  // TExportListView Units
  EC_Strings;

constructor TExportListView.Create(AOwner: TComponent);
var
  i: Integer;
begin
  inherited;

  // Detect First ListView on Form (if there's not already a listview assigned)
  if not Assigned(FListView) then
  with TForm(Owner) do
    for i := 0 to ComponentCount - 1 do
      if Components[i] is TListView then
      begin
        FListView := TListView(Components[i]);
        break;
      end;
end;

procedure TExportListView.InitExport;
begin
  inherited;

  FCount.SelectedItems := ListView.SelCount;
end;

procedure TExportListView.CleanUpExport;
begin
  inherited;

  // Free Column Order Array
  if Assigned(ColOrderArray) then
    FreeMem(ColOrderArray);
end;

function TExportListView.GetNumExportableItems: Integer;
begin
  with FListView do
    if (SelectedRowsOnly in FOptions) and (SelCount > 0) then
      Result := SelCount
    else Result := Items.Count;
end;

function TExportListView.CurrentItemSelected: Boolean;
begin
  Result := FListView.Items[FCount.CurrentRow].Selected;
end;

function TExportListView.ColAllowed(ListCol: TListColumn): Boolean; // Returns True if the Column should be exported
begin
  Result := True; // Assume Success

  // Don't Export if it's invisible and we're not allowed to export invisible columns
  if (ListCol.Width = 0) and not (ExportInvisibleCols in FOptions) then
    Result := False;
end;

function TExportListView.GetColIndex(aIndex: Integer): Integer;
begin
  // Get Column to Work On, but in order that they are displayed on the Screen, not the
  // internal order that delphi tracks in the TListColumns object.
  if Assigned(ColOrderArray) then
    Result := ColOrderArray^[aIndex] // Return Result from Column Order Array,
  else Result := aIndex;             // unless it doesn't exist.
end;

function TExportListView.GetColValue(aIndex: Integer): String;
begin
  Result := FListView.Column[aIndex].Caption;
end;

function TExportListView.GetExportColumns(var aColumns: TStringList): TStringList;
  function GetListViewColumnOrder: Pointer;
  const
    LVM_GetListViewColumnOrderARRAY = LVM_FIRST + 59;
  begin
    with FListView do
    try
      // Allocate Memory for Columns Order Array
      GetMem(Result, Columns.Count * SizeOf(Integer));

      // Get Order of Columns
      if SendMessage(Handle, LVM_GetListViewColumnOrderARRAY, Columns.Count, LPARAM(Result)) = 0 then
      begin
        // On failure, free memory and return nil
        Result := nil;
        if Assigned(Result) then
          FreeMem(Result);
      end;

    except
      Result := nil;
    end;
  end;

var
  i: Integer;

begin
  Result := inherited GetExportColumns(aColumns);

  with FListView do
  begin
    // Allocate StringList Capacity
    {$IFDEF Delphi3Up}
    aColumns.Capacity := Columns.Count;
    {$ENDIF}

    ColOrderArray := GetListViewColumnOrder; // Get the Order of the columns (it may not be as listed in the
                                             // TListColumns object because a user may have dragged them around

    for i := 0 to Columns.Count - 1 do                  // For Each Column...
      if ColAllowed(Column[GetColIndex(i)]) then        // ...that we're allowed to export...
        aColumns.Add(GetColumnCaption(GetColIndex(i))); // ...Pass it on to the Descendant.
  end;
end;

procedure TExportListView.GetExportItemData(var aExportItem: TStringList);
var
 j, k: Integer;
begin
  // Allocate StringList Capacity
  {$IFDEF Delphi3Up}
  aExportItem.Capacity := FExportColumns.Count;
  {$ENDIF}

  // All Fields in Item's Row for Columns that are Visible -> S
  with FListView do
    for j := 0 to Columns.Count - 1 do
    begin
      k := GetColIndex(j); // Get Index of Column We're Dealing With

      if ColAllowed(Column[k]) then // As long as the column is allowed to be exported
        if k = 0 then                             // First Column get's special treatment
          aExportItem.Add(Items[FCount.CurrentRow].Caption)       // (Use TListItem.Caption property
        else
          if Items[FCount.CurrentRow].SubItems.Count > j - 1 then    // If this SubItem exists,
            aExportItem.Add(Items[FCount.CurrentRow].SubItems[k-1]);  // we add from Sub Items
    end;
end;

function TExportListView.HasData: Boolean;
begin
  Result := True; // Assume Success

  // Abort under the following circumstances:
  if Assigned(FListView) then
  with FListView do
  begin
    if Items.Count = 0 then // If No Items to Export
    begin
      MessageDlg(cNoDataError, mtWarning, [mbOk], 0);
      Result := False;
      Exit;
    end
    else
      if (SelectedRowsOnly in FOptions) and
         (not (ExportAllWhenNoneSelected in FOptions)) and
         (SelCount = 0) then
      begin
        MessageDlg(cNoneSelectedError, mtWarning, [mbOk], 0);
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
  else raise Exception.CreateFmt(cNoComponentError, ['TListView']);
end;

// *** Set Access Methods ***

procedure TExportListView.SetListView(Value: TListView);
begin
  if Value <> FListView then
  begin
    FListView := Value;
    if Assigned(Value) then
      FListView.FreeNotification(Self); // Add Notification if the ListView is Destroyed
  end;
end;

procedure TExportListView.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited;

  // Assign nil to FListView when the Assigned ListView is Freed
  if (AComponent = FListView) and (Operation = opRemove) then
    FListView := nil;
end;

end.
