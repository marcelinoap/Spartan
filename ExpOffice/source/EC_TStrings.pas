unit EC_TStrings;

{$I CompConditionals.Inc}

interface

uses Classes, Grids,

     // TExportStrings Units
     EC_Main;

type
  PStrings = ^TStrings;

  TSelCountEvent = procedure(var SelCount: Integer) of object;
  TSelectedEvent = procedure(i: Integer; var Selected: Boolean) of object;

  TExportStrings = class(TExportXComponent)
  private
    FStrings: TStrings;
    FExportStrings, ExternalStrings: ^TStrings;
    FUseExternalStrings: Boolean;

    FOnSelCount: TSelCountEvent;
    FOnSelected: TSelectedEvent;

    // *** Set Access Methods ***
    procedure SetStrings(Value: TStrings);

    // *** Export Procedures ***
    procedure CleanUpExport; override;

    function HasData: Boolean; override;

    function GetDefaultColWidth(ColIndex: Integer): Integer; override;    
    function GetNumExportableItems: Integer; override;
    function CurrentItemSelected: Boolean; override;

    function GetColValue(aIndex: Integer): String; override;
    function GetExportColumns(var aColumns: TStringList): TStringList; override;
    procedure GetExportItemData(var aExportItem: TStringList); override;

  public
    procedure UseStrings(pS: PStrings);

    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

  published
    property Strings: TStrings read FStrings write SetStrings;

    // Events
    property OnSelCount: TSelCountEvent read FOnSelCount write FOnSelCount;
    property OnSelected: TSelectedEvent read FOnSelected write FOnSelected;
  end;

implementation

uses
  CommCtrl, Windows, SysUtils, Dialogs, Forms,

  // TExportStrings Units
  EC_Strings;

const
  cStringsColWidth = 70;

constructor TExportStrings.Create(AOwner: TComponent);
begin
  inherited;

  // Set Illegal Options
  FIllegalOptions := [ExportInvisibleCols];

  // Create Objects
  FStrings := TStringList.Create;
end;

destructor TExportStrings.Destroy;
begin
  // Free Objects
  FStrings.Free;
end;

procedure TExportStrings.CleanUpExport;
begin
  inherited;

  FUseExternalStrings := False;
end;

function TExportStrings.GetDefaultColWidth(ColIndex: Integer): Integer;
begin
  Result := cStringsColWidth;
end;

function TExportStrings.GetNumExportableItems: Integer;
var
  i: Integer;
begin
  // Calculate Number of Selected Items
  if (SelectedRowsOnly in FOptions) and Assigned(FOnSelCount) then
    FOnSelCount(FCount.SelectedItems) // Get Number of Selected Items
  else FCount.SelectedItems := 0;

  // Return Number of Exportable Items
  if (SelectedRowsOnly in FOptions) and not
     ((ExportAllWhenNoneSelected in FOptions) and (FCount.SelectedItems = 0)) then
    Result := FCount.SelectedItems
  else
  begin
    Result := FExportStrings^.Count;

    // Don't Export Blank Strings at End of TStrings
    for i := Result - 1 downto 0 do
    begin
      if Length(Trim(FExportStrings^[i])) > 0 then
        break
      else Dec(Result);
    end;
  end;
end;

function TExportStrings.CurrentItemSelected: Boolean;
begin
  // Get from event
  if Assigned(FOnSelected) then
    FOnSelected(FCount.CurrentRow, Result);
end;

function TExportStrings.GetColValue(aIndex: Integer): String;
begin
  Result := ''; // Force it to get from Captions property.
end;

function TExportStrings.GetExportColumns(var aColumns: TStringList): TStringList;
begin
  Result := inherited GetExportColumns(aColumns);

  // Allocate StringList Capacity
  {$IFDEF Delphi3Up}
  aColumns.Capacity := 1;
  {$ENDIF}

  aColumns.Add(GetColumnCaption(0));  // ...Pass it on to the Descendant.
end;

procedure TExportStrings.GetExportItemData(var aExportItem: TStringList);
begin
  // Allocate StringList Capacity
  {$IFDEF Delphi3Up}
  aExportItem.Capacity := 1;
  {$ENDIF}

  // Add
  aExportItem.Add(FExportStrings^[FCount.CurrentRow]);
end;

function TExportStrings.HasData: Boolean;
begin
  Result := True; // Assume Success

  // Set Pointer
  if FUseExternalStrings then
    FExportStrings := Pointer(ExternalStrings)
  else FExportStrings := @FStrings;

  // Abort under the following circumstances:
  if Assigned(FExportStrings^) then
  with FExportStrings^ do
  begin
    if Count = 0 then // If No Items to Export
    begin
      MessageDlg(cNoDataError, mtWarning, [mbOk], 0);
      Result := False;
      Exit;
    end
    else
      if (SelectedRowsOnly in FOptions) and
         (not (ExportAllWhenNoneSelected in FOptions)) and
         (GetNumExportableItems = 0) then
      begin
        MessageDlg(cNoneSelectedError, mtWarning, [mbOk], 0);
        Result := False;
        Exit;
      end;
  end
  else raise Exception.CreateFmt(cNoComponentError, ['TStrings']);
end;

// *** Set Access Methods ***

procedure TExportStrings.SetStrings(Value: TStrings);
begin
  if Assigned(Value) then
    FStrings.Assign(Value);
end;

procedure TExportStrings.UseStrings(pS: PStrings);
begin
  FUseExternalStrings := True;
  ExternalStrings     := Pointer(pS);
end;

end.
