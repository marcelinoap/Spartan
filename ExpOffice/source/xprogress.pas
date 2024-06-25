unit XProgress;

{$I CompConditionals.inc} // Component Conditional Defines

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, ExtCtrls;

type
  TProgressForm = class(TForm)
    ProgressBar: TProgressBar;
    ExportPanel: TPanel;
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    StartTime: dWord;

    ProgressUpdate: Integer;
    MaxUnknown: Boolean;
    ExportStatusMsg: String;

  public
    function UpdateProgress(CurrentPos: Integer): Boolean;
  end;

function CreateProgress(aCaption, aInitialMsg, aExportStatusMsg: String; aCount: Integer): TProgressForm;

implementation

{$R *.DFM}

// *** Progress Methods ***
function CreateProgress(aCaption, aInitialMsg, aExportStatusMsg: String; aCount: Integer): TProgressForm;
const
  cNumIncrements = 25;
  cRunningCountFormHeight = 57;
var
  ProgressForm: TProgressForm;
begin
  ProgressForm := TProgressForm.Create(nil);
  with ProgressForm do
  begin
    ExportPanel.Caption := aInitialMsg; // Set Initial Message
    StartTime := GetTickCount; // Get Start Time

    MaxUnknown := aCount < 0; // If We don't know the max value, we have to do a running count

    Caption := aCaption; // Set Caption
    ExportStatusMsg := aExportStatusMsg; // Set Export Status Message
    ProgressUpdate := aCount div cNumIncrements + 1; // otherwise update every so often

    // ProgressBar Initialization Code
    with ProgressBar do
      if MaxUnknown then
      begin
        Visible := False;
        ProgressForm.Height := cRunningCountFormHeight; // Adjust Form Height
        ExportPanel.Visible := True;
      end
      else Max := aCount div ProgressUpdate; // Set Progress Range
  end;

  Result := ProgressForm;
end;

// Returns False if user cancels progress dialog
function TProgressForm.UpdateProgress(CurrentPos: Integer): Boolean;
var
  ElapsedTime: Single;
begin
  Result := True;

  if CurrentPos mod ProgressUpdate = 0 then
    if MaxUnknown then
    begin
      // Calculate Elapsed Time
      ElapsedTime := (GetTickCount - StartTime) / 1000;

      // Display Status Message
      ExportPanel.Caption := Format(ExportStatusMsg, [CurrentPos, ElapsedTime]);
      Application.ProcessMessages;
    end
    else
      with ProgressBar do
      begin
        Application.ProcessMessages;

        if Position + Step > Max then
          Position := Max
        else StepIt;

        if not Visible then Result := False;
      end;
end;

// *** Delphi IDE Maintained ***

procedure TProgressForm.FormCreate(Sender: TObject);
begin
{$IFDEF Delphi4Up}
  ProgressBar.Smooth := True;
{$ENDIF}
end;

procedure TProgressForm.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_Escape then
    Close;
end;

end.
