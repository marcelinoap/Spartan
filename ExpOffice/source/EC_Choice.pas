unit EC_Choice;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls,

  // TExportListView Units
  EC_Main;

type
  TExportChoiceForm = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    GroupBox1: TGroupBox;
    ExportTypeComboBox: TComboBox;
    ViewRadioButton: TRadioButton;
    FileRadioButton: TRadioButton;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
    procedure ViewRadioButtonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure ExportTypeComboBoxChange(Sender: TObject);

  public
    NumExportItems: Integer;
    ViewOnlyChoosen: Boolean;
    SelectedExportType: TExportType;
  end;

var
  ExportChoiceForm: TExportChoiceForm;

implementation

{$R *.DFM}

uses EC_Strings, ShellAPI;

procedure TExportChoiceForm.FormCreate(Sender: TObject);
begin
  // Set Captions
  OkBtn.Caption           := cChooseOkCaption;
  CancelBtn.Caption       := cChooseCancelCaption;
  Groupbox1.Caption       := cChooseInstructions + ' ';
  ViewRadioButton.Caption := cChooseScreenCaption;
  FileRadioButton.Caption := cChooseFileCaption;

  // Initialize
  ViewOnlyChoosen := True;
end;

procedure TExportChoiceForm.FormShow(Sender: TObject);
var
  i: Integer;
begin
  // Set Form Caption
  if NumExportItems > 0 then
    Caption := Format(cChooseDialogCaption, [NumExportItems])
  else Caption := cChooseDialogCaption2; // If # items unknown, don't show the count then

  // Select Most-Recently-Used Export Type
  with ExportTypeComboBox do
  begin
    for i := 0 to Items.Count - 1 do
      if Items[i] = cExportTypes[SelectedExportType] then
      begin
        ItemIndex := i;
        break;
      end;

    // If default Export Type is not an Allowed Type,
    if ItemIndex = -1 then
      ItemIndex := 0; // ... default to first Allowed Type
  end;

  // Update Radio Buttons
  ExportTypeComboBoxChange(Sender);
end;


procedure TExportChoiceForm.ViewRadioButtonClick(Sender: TObject);
begin
  // Select Exporting To File, or Viewing on Screen
  ViewOnlyChoosen := Sender = ViewRadioButton;
end;

procedure TExportChoiceForm.OKBtnClick(Sender: TObject);
begin
  with ExportTypeComboBox do
  begin
    SelectedExportType := GetExportType(Items[ItemIndex]);

//Jair  -> selecionando excel abre csv
//    if SelectedExportType=xMicrosoft_Excel then
//       SelectedExportType := GetExportType(Items[5]);

  end;
end;

procedure TExportChoiceForm.ExportTypeComboBoxChange(Sender: TObject);
type
  TIconInfo = record
    Executable: String;
    Index: Integer;
  end;

  function GetAssociatedIcon(aExt: String): TIconInfo;
  const
    cIconKey = 'DefaultIcon';
    cOpenKey  = 'open';
  var
    S: String;
  begin
    Result.Executable := ''; // Assume Failure

    with TMyReg.Create do
    try
      RootKey := HKEY_CLASSES_ROOT;

      if OK_ReadOnly('.' + aExt) then // Attempt to Open Extension Association key
      begin
        S := ReadString(''); // Get Class Name
        CloseKey;

        if OK_ReadOnly(S) and OK_ReadOnly(cIconKey) then
        begin
          S := ReadString(''); // Get Filename
          Result.Executable := Copy(S, 1, Pos(',', S) - 1); // Get Executable
          Result.Index      := StrToInt(Copy(S, Pos(',', S) + 1, Length(S))); // Get Index
        end;

      end;

    finally
      CloseKey;
      Free;
    end;
  end;

  function GetWindowsDir: String;
  // Returns '' on failure
  const
    MaxPathSize = 1024;
  var
    Buff: array[1..MaxPathSize] of Char;
  begin
    if GetWindowsDirectory(@Buff, MaxPathSize) <> 0 then
      Result := pchar(@Buff)
    else Result := '';
  end;

  procedure ClearIcon;
  begin
    with Image1.Picture do
    begin
      if Assigned(Icon) then
        DestroyIcon(Icon.Handle);
      Icon := nil;     // Clear Icon Display
    end;
  end;

var
  NewExportType: TExportType;
  IconInfo: TIconInfo;
begin
  // Figure out what the New Export Type is
  with ExportTypeComboBox do
    NewExportType := GetExportType(Items[ItemIndex]);

  // Enable Disable Radio Buttons based on what the new export type is
  FileRadioButton.Enabled := not (NewExportType = xClipboard); // Can't Export Clipboard to file, but all others can
  ViewRadioButton.Enabled := TExportXComponent(Owner).ExportAppInstalled(NewExportType) and // Can View only if Viewer Installed
                             FileRadioButton.Enabled;

  if ViewRadioButton.Enabled or (NewExportType = xClipboard) then // If Viewing not allowed
  begin
    // Make Sure Viewing is the default choice
    ViewRadioButton.Checked := True;

    // Get Associated File containing Shell Icon
    if NewExportType = xClipboard then // If clipboard, get from clipbrd.exe
    begin
      IconInfo.Executable := GetWindowsDir + '\clipbrd.exe';
      IconInfo.Index      := 0;
    end
    else IconInfo := GetAssociatedIcon(cExportTypeExtensions[NewExportType]); // otherwise, get the associted shell icon for the export type extension

    // Display Shell Icon
    if Length(IconInfo.Executable) > 0 then
    begin
      ClearIcon; // Clear Icon
      Image1.Picture.Icon.Handle := ExtractIcon(0, PChar(IconInfo.Executable), IconInfo.Index);
      if Image1.Picture.Icon.Handle <= 1 then // Get Icon from Executable
        ClearIcon;
    end
    else ClearIcon;
  end
  else // If Viewing not allowed
  begin
    with FileRadioButton do
    begin
      Checked := True;                      // Check File Radio Button (this ensures view radio button is not checked)
      if not Enabled then Checked := False; // If not enabled, we don't want any buttons checked
    end;
    ClearIcon;     // Clear Icon Display
  end
end;

end.
