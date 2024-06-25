unit ExportExtLVReg;

interface

uses EC_ExtLView, EC_EnhLView;

const
  {$I CompConstants.inc} // Component Constants & Conditional Defines

procedure Register;

implementation

uses Classes;

procedure Register;
begin
  RegisterComponents(cExportTab, [TExportExtListView, TExportEnhListView]);
end;

end.
