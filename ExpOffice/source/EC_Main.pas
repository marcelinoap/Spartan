{------------------------------------------------------------------------------}
{ License Agreement & Disclaimer                                               }
{------------------------------------------------------------------------------}
{ You may not distribute any files for this component except the unregistered  }
{ version's .Zip File. You may definately not distribute the source, it is for }
{ your eyes only. Anyone else wishing to see the source must purchase it at    }
{ http://www.igather.com/components                                            }
{                                                                              }
{ The unregistered version's .Zip file can be included on any kind of media or }
{ distributed through any kind of medium.                                      }
{                                                                              }
{ Although this component has been thoroughly tested and documented, neither   }
{ Y-Tech Corporation nor any of it's employees nor the author will be held     }
{ responsible for any damage arising from it's use or misuse.                  }
{------------------------------------------------------------------------------}

unit EC_Main;

{$include CompConditionals.Inc}
//{$include reg.inc}

interface

uses Classes, ComCtrls, Windows, Dialogs, Registry, Graphics,

     // TExportXComponent Units
     XProgress, EC_Strings, jjlib;

const
  {$I CompConstants.inc} // Component Constants

type
  TExportType = (xHTML,
                 xMicrosoft_Word,
                 xMicrosoft_Excel,
                 xText,
                 xRichText,
                 xText_Comma_Delimited,
                 xText_Tab_Delimited,
                 xDIF,
                 xSYLK,
                 xClipboard);

  TExportTypeSet = set of TExportType;

  TExportTypeConstants = array [Low(TExportType)..
                                High(TExportType)] of String;

  TExportOptions = set of (DetailedMode,
                           ExportAllWhenNoneSelected,
                           ExportInvisibleCols,
                           NumberRows,
                           SelectedRowsOnly,
                           ShowColHeaders,
                           ShowClipboardMsg,
                           ShowProgress,
                           TimeStamp,
                           TimeDisplaySeconds,
                           Time24HourFormat);

  THTML_Options = set of (htShowGridLines,
                          htColHeadersBold,
                          htAutoLink,
                          htOddRowColoring,
                          htDisplayTitle,
                          htHorzRules,
                          htIgnoreLineBreaks);

  TExportCount = record
    CurrentRow,             // Current Row # including rows not exported due to them not being selected and such (ie. Offset)
    RowsExported,           // Current Row # not including rows not exported (ie. num rows exported so far)
    ExportItems,            // Total Number of Rows being Exported (not always known)
    SelectedItems: Integer; // Number of Items Selected - must be set by descendant
  end;

const
  // ExportType Constants
  cExportTypes: TExportTypeConstants =
    (ccHTML,
     ccMicrosoftWord,
     ccMicrosoftExcel,
     ccText,
     ccRichText,
     ccCSVText,
     ccTabText,
     ccDIF,
     ccSYLK,
     ccClipboard);

  cExportTypeExtensions: TExportTypeConstants =
    ('htm',
     'doc',
     'xls',
     'txt',
     'rtf',
     'csv',
     'txt',
     'dif',
     'slk',
     ''); // Can't Export Clipboard to File - so it's blank.

  // Default Options
  cDefOptions: TExportOptions      = [NumberRows, ShowColHeaders, ShowClipboardMsg, ShowProgress, TimeStamp];
  cDefHTML_Options : THTML_Options = [htAutoLink, htColHeadersBold, htOddRowColoring];
  cDefaultTruncateSymbol           = '-';

  // RichText Constants (more further down at the next "const" section)
  cRTF_DefaultLeftRightMargin = 1250; // In TWIPS
  cRTF_DefaultTopBottomMargin = 1000; // In TWIPS

  // Max Col Widths Dynamic Array Stuff
  cMaxColWidthsArraySize = (65520 div SizeOf(Integer)); // Should be big enough ;)

  // File Dialog Options & Titles for Export File Dialog
  cFileDialogOptions = [ofFileMustExist, ofHideReadOnly, ofPathMustExist,
                        ofOverWritePrompt];

type
  //----------------------------------------------------------------------------
  // Class        : TAutoCleanupStringList (TStringList)
  //----------------------------------------------------------------------------
  // Purpose      : Automatically frees associated objects on Clear/Free.
  // Functionality: Use Clear as Normal, but call AutoFree instead of Free.
  //----------------------------------------------------------------------------

  TAutoCleanupStringList = class(TStringList)
  public
    procedure Clear; override;
    procedure AutoFree;
  end;

  //----------------------------------------------------------------------------
  // Class        : TMyReg (TRegistry)
  //----------------------------------------------------------------------------
  // Purpose      : It's sole purpose is to give OpenKeyReadOnly capabilities to
  //                D2/D3.
  // Functionality: Adds OK_ReadOnly method which is a copy of D4's
  //                OpenKeyReadOnly method.
  // Conditions   : You should never use the OpenKeyReadOnly method, always use
  //                OK_ReadOnly - it does the same thing but works with D2/D3
  //                also.
  //----------------------------------------------------------------------------

  TMyReg = class(TRegistry)
  public
    function OK_ReadOnly(const Key: String): Boolean;
  end;

  //----------------------------------------------------------------------------
  // Class        : TExportXComponent (TComponent)
  //----------------------------------------------------------------------------
  // Purpose      : Base class for all TExportX Components
  // Functionality: Abstract Class.
  //----------------------------------------------------------------------------

  // The HTML Template Types
  THTML_Template = (htClassic,
                    htSimple,
                    htPlain,
                    htMurky,
                    htMS_Money,
                    htColorful,
                    htGray,
                    htOlive,
                    htBW);

  THTML_Templates = array [Low(THTML_Template)..
                           High(THTML_Template)] of String;


  // TExportColumnsInfo Types
  TColType = (ctUnknown,
              ctString,
              ctMemo,
              ctFmtMemo,
              ctNumber,
              ctCurrency,
              ctBoolean,
              ctDate,
              ctTime,
              ctDateTime,
              ctDateTime_ShowDateOnly,
              ctDateTime_ShowTimeOnly);

{  TColFormat = record // Deferred until next version
    Alignment: TAlignment;
    Font: TFont;
  end;}

{  TNumberStyle = (nsThousandSeparators, // Display Thousands Separators
                  nsScientificNotation, // Display using Scientific Notation
                  nsPercentage);        // Display as a Percentage}

  TExportColumnInfo = class(TObject)
  public
    ColType: TColType;               // Data Type of the Data in the Column
//    NumberStyle: TNumberStyle;   // Number Style for all Data in Column (applies only to
                                 // numeric ColTypes)
//    ColFormat,                   // Formatting Info for the Data in the Column
//    ColHeaderFormat: TColFormat; // Formatting Info for the Column Header
  end;

  TBooleanValue = (bvNull, bvFalse, bvTrue);

  TExportItemInfo = class(TObject)
  public
    UseDisplayString: Boolean;
    DisplayString: String;

    BooleanValue: TBooleanValue; // Used Only if ColType = ctBoolean
  end;

  TPrint_Options = class(TPersistent)
  private
    FDefaultFont,
    FDetailedTitle,
    FDetailedHeader,
    FDetailedFooter,
    FDetailedFont: TFont;

  public
    constructor Create;
    destructor Destroy; override;

  published
    property DefaultFont    : TFont read FDefaultFont    write FDefaultFont;
    property DetailedTitle  : TFont read FDetailedTitle  write FDetailedTitle;
    property DetailedHeader : TFont read FDetailedHeader write FDetailedHeader;
    property DetailedFooter : TFont read FDetailedFooter write FDetailedFooter;
    property DetailedFont   : TFont read FDetailedFont   write FDetailedFont;
  end;

  TXL_Options = class(TPersistent)
  private
    FPrintGridlines,
    FPrintHeadings: Boolean;
    FPageHeader,
    FPageFooter: String;

  public
    constructor Create;

  published
    property PrintGridlines: Boolean read FPrintGridlines write FPrintGridlines;
    property PrintHeadings : Boolean read FPrintHeadings  write FPrintHeadings;
    property PageHeader    : String  read FPageHeader     write FPageHeader;
    property PageFooter    : String  read FPageFooter     write FPageFooter;
  end;

  TRTF_Options = class(TPersistent)
  private
    FLeftMargin,
    FRightMargin,
    FTopMargin,
    FBottomMargin: Word;

  public
    constructor Create;

  published
    property LeftMargin  : Word read FLeftMargin   write FLeftMargin   default cRTF_DefaultLeftRightMargin;
    property RightMargin : Word read FRightMargin  write FRightMargin  default cRTF_DefaultLeftRightMargin;
    property TopMargin   : Word read FTopMargin    write FTopMargin    default cRTF_DefaultTopBottomMargin;
    property BottomMargin: Word read FBottomMargin write FBottomMargin default cRTF_DefaultTopBottomMargin;
  end;

  TWriteFooterEvent = procedure (Sender: TObject; NumItemsExported: Integer) of object;

  TExportXComponent = class(TComponent)
  private
    procedure OnActivate(Sender: TObject); // Hook Form's OnActivate Event

  protected
    FOldActivate: TNotifyEvent; // Store Old Application.OnActivate Event
    FAppMinimized: Boolean;     // If True the App Has been minimized by the Export Object

    FDefaultColWidthSpacing: Integer; // See .GetDefaultColWidth

    // *** Options ***
    FOptions,
    FIllegalOptions: TExportOptions;
    FHTML_Options: THTML_Options;


    // *** Objects ***
    FExportColumns,
    FFormatFooter: TStringList; // The Actual Footer Text to be Written (after number of rows exported has been inserted)
    FCount: TExportCount;

    FPrint_Options: TPrint_Options;
    FXL_Options: TXL_Options;
    FRTF_Options: TRTF_Options;

    // *** Property Variables ***
    FTitle,
    FExportFile: String;

    FViewOnly: Boolean;

    FExportType: TExportType;
    FAllowedTypes: TExportTypeSet;
    FCaptions,
    FColWidths,
    FHeader,
    FFooter: TStrings;

    FHTML_CustomTemplate: String;
    FHTML_Template: THTML_Template;

    FTruncateSymbol: String;

    FLastExportType: TExportType;

    // *** Event Variables ***
    FOnBeginExport,
    FOnExportFinished,
    FOnExportFailed,
    FOnPrintFailed: TNotifyEvent;

    FOnWriteFooter: TWriteFooterEvent;

    // *** Set Access Methods ***
    procedure SetCaptions(Value: TStrings);
    procedure SetColWidths(Value: TStrings);
    procedure SetPrint_Options(Value: TPrint_Options);
    procedure SetXL_Options(Value: TXL_Options);
    procedure SetRTF_Options(Value: TRTF_Options);

    procedure SetTitle(Value: String);
    procedure SetHeader(Value: TStrings);
    procedure SetFooter(Value: TStrings);
    procedure SetTruncateSymbol(Value: String);

    procedure SetOptions(Value: TExportOptions);
    procedure SetHTML_Options(Value: THTML_Options);

    procedure SetHTML_CustomTemplate(Value: String);

    // *** Misc. procedures ***
    procedure Loaded; override;

    function Strip(S: String): String;
    procedure TrimTrailingBlankStrings(var S: TStrings);

    function GetDefaultColWidth(ColIndex: Integer): Integer; virtual;

    // *** procedures which must be overriden by Descendants ***
    procedure InitExport; virtual;
    procedure CleanUpExport; virtual;

    function GetColumnCaption(aIndex: Integer): String;
    function GetColValue(aIndex: Integer): String; virtual; abstract;

    function GetNumExportableItems: Integer; virtual; abstract;
    function CurrentItemSelected: Boolean; virtual;
    function GetExportColumns(var aColumns: TStringList): TStringList; virtual;
    procedure InitExportItems; virtual;
    procedure SkipToNextItem; virtual;
    procedure GetNextExportItem(var aExportItem: TStringList); virtual;
    procedure GetExportItemData(var aExportItem: TStringList); virtual; abstract;
    function MoreExportItems: Boolean; virtual;

    function HasExportableColumns: Boolean;
    function HasData: Boolean; virtual; abstract;

    // Properties
    property ColWidths: TStrings read FColWidths write SetColWidths;

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function Execute: Boolean;
    function Print: Boolean;
    function Choose: Boolean;

    procedure RestoreApp; // Restore App if it has been minimized by the Export Object
    procedure SetTempDir(TempDir: String); // User can set the Temp Dir himself

    function ExportAppInstalled(ExportType: TExportType): Boolean; // Undocumented
    procedure PopulateStrings(S: TStrings); // For Backwards Compatibility Only

  published
    property Options: TExportOptions read FOptions write SetOptions;
    property HTML_Options: THTML_Options read FHTML_Options write SetHTML_Options;

    property ExportFile: String read FExportFile write FExportFile;

    property Title : String read FTitle  write SetTitle;
    property Header: TStrings read FHeader write SetHeader;
    property Footer: TStrings read FFooter write SetFooter;

    property ViewOnly: Boolean read FViewOnly write FViewOnly default True;

    property ExportType: TExportType read FExportType write FExportType;
    property AllowedTypes: TExportTypeSet read FAllowedTypes write FAllowedTypes;
    property Captions: TStrings read FCaptions write SetCaptions;

    property HTML_CustomTemplate: String read FHTML_CustomTemplate write SetHTML_CustomTemplate;
    property HTML_Template: THTML_Template read FHTML_Template write FHTML_Template default htSimple;

    property TruncateSymbol: String read FTruncateSymbol write SetTruncateSymbol;

    property Print_Options: TPrint_Options read FPrint_Options write FPrint_Options;
    property XL_Options: TXL_Options read FXL_Options write SetXL_Options;
    property Word_RTF_Options: TRTF_Options read FRTF_Options write SetRTF_Options;

    property OnBeginExport: TNotifyEvent read FOnBeginExport write FOnBeginExport;
    property OnExportFinished: TNotifyEvent read FOnExportFinished write FOnExportFinished;
    property OnExportFailed: TNotifyEvent read FOnExportFailed write FOnExportFailed;
    property OnPrintFailed: TNotifyEvent read FOnPrintFailed write FOnPrintFailed;

    property OnWriteFooter: TWriteFooterEvent read FOnWriteFooter write FOnWriteFooter;    
  end;

  TSharewareExportComponent = class(TExportXComponent)
  public
    constructor Create(AOwner: TComponent); override;
  end;

  // This array stores the maximum width of each column
  TMaxColWidths = array[0..cMaxColWidthsArraySize - 1] of Integer;

  //----------------------------------------------------------------------------
  // Class        : TCustomExport (TObject)
  //----------------------------------------------------------------------------
  // Purpose      : Base class for all Export (Format) Objects including
  //                Printer and Clipboard Exports.
  // Functionality: Abstract Class. 
  //----------------------------------------------------------------------------

  TCustomExport = class(TObject)
  private
    FViewerHandle: HWnd;
    FOwner: TExportXComponent;
    FColumns: TStringList;

    PMaxColWidths: ^TMaxColWidths; // Pointer to MaxColWidths array

    FNumEntriesExported: Integer; // Sometimes this is different than the value of the "Entry" paramter.
                                  // If ShowColHeaders is True, then it should be one larger with xDIF, CSV, xSYLK, TabText, etc... formats
    FExportType: TExportType;
    FExportFile,        // File name to Export To
    FExtension: String;

    FUserCancelled,      // If True, User cancelled so we don't want to do anything anymore
    FExportSuccessful,
    FExportToFile: Boolean; // If false, exports to screen only

    procedure SetInitialMaxColWidths(var S: TStringList);
    procedure UpdateMaxColWidths(var S: TStringList);

  protected
    // Variables
    FCreatedSuccessfully: Boolean;
    FProgForm: TProgressForm;

    // *** Export Data Variables ***
    FUnderlineLength,
    FMaxColumnLength: Integer;

    // *** Read-Only Methods ***
    property ExportSuccessful: Boolean read FExportSuccessful;
    property ExportToFile: Boolean read FExportToFile;

    // Padded Export Methods
    function Pad(S: String; aLength: Integer; RightAlignNumbers: Boolean): String;
    function GetPaddedHeader(var Columns: TStringList): String;
    function GetPaddedEntry(var Items: TStringList; Entry: Integer): String;
    function GetPaddedFooter: String;

    // Strip Methods
    function StripConvert(S, sTab, sNewLine: String): String;
    function StripInvalidChars(S, ValidChars: String): String;
    function StripNumberFormat(S, RawNumber: String): String; 

    // *** Protected Methods ***
    procedure CalculateMaxColumnLength;

    function GetRoundedNumber(RawNumber: String; ItemInfo: TExportItemInfo): String;

    function GetDisplayString(var aItems: TStringList; aIndex: Integer): String;

    function GetDateFormatString: String;
    function GetTimeFormatString: String;
    function GetDateTimeFormatString: String;
    function GetCurrencyFormatString: String;

    function ConvertToDate(S: String): TDateTime;
    function ConvertToDateTime(S: String): TDateTime;

    function Decimalize(S: String): String;
    function IsNumber(S: String): Boolean;
//    function IsInteger(S: String): Boolean; - maybe useful in some future version...

    function GetColInfo(Col: Integer): TExportColumnInfo;
    function GetColType(i: Integer; StringValue: String): TColType;
    function GetColWidth(i: Integer): Integer;

    function GetTempFile(FileExt: String): String;
    function GetTimeStamp: String;

    procedure WriteHeader(var Columns: TStringList); virtual; // The Descendant Object must override this place
    procedure WriteEntry(var Items: TStringList; Entry: Integer); virtual;
    procedure WriteFooter; virtual; // This one is optional, you don't have to override it

    procedure HandleError(Exception: TObject); virtual; // This is optional too.

  public
    // *** Public Properties ***
    property ExportFile: String read FExportFile;

    // *** Public Methods ***
    constructor Create(aOwner: TExportXComponent; aExportFile: String; aExportType: TExportType); virtual;
    destructor Destroy; override;
    function DoExport: Boolean; // Perform Export
  end;

  //----------------------------------------------------------------------------
  // Class        : TCustomShellViewExport
  //----------------------------------------------------------------------------
  // Purpose      : Base class for all Export objects that show the files to the
  //                user using ShellExecute
  // Functionality: Abstract class.
  //----------------------------------------------------------------------------

  TCustomShellViewExport = class(TCustomExport)
  private
    FOutputFile: String; // Might be different from Export File because if not exporting to file, we have to create temporary one

  protected
    ShellViewErrorMsg: String;  // Assign an Error Message to this if you want a different one than the default file error message

    property OutputFile: String read FOutputFile; // Read-Only
  public
    constructor Create(aOwner: TExportXComponent; aExportFile: String; aExportType: TExportType); override;
    destructor Destroy; override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TCustomTextExport
  //----------------------------------------------------------------------------
  // Purpose      : Base class for all Text File Export objects.
  // Functionality: Abstract class.
  //----------------------------------------------------------------------------

  TCustomTextExport = class(TCustomShellViewExport)
  private
    F: TextFile;
    FAlreadyClosed: Boolean;

  protected
    procedure CloseTextFile;
    procedure HandleError(Exception: TObject); override;

  public
    constructor Create(aOwner: TExportXComponent; aExportFile: String; aExportType: TExportType); override;
    destructor Destroy; override;
  end;

  // *** Export Classes ***

  //----------------------------------------------------------------------------
  // Class        : TClipboardExport
  //----------------------------------------------------------------------------
  // Purpose      : Exports data to the Clipboard in a Grid format padded with
  //                spaces.
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  // Uses         : Uses the ShowClipboardMsg Option to determine if it should
  //                display a dialog to the user indicating that the export has
  //                completed.
  // Notes        : If a column is not as wide as the data, the data will get
  //                truncated.
  //----------------------------------------------------------------------------

  TClipboardExport = class(TCustomExport)
  private
    FClipboardText: String;

  protected
    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
    procedure WriteFooter; override;

  public
    destructor Destroy; override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TDetailedClipboardExport
  //----------------------------------------------------------------------------
  // Purpose      : Exports data to the Clipboard in Detailed Mode.
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  // Uses         : Uses the ShowClipboardMsg Option to determine if it should
  //                display a dialog to the user indicating that the export has
  //                completed.
  //----------------------------------------------------------------------------

  TDetailedClipboardExport = class(TCustomExport)
  private
    S: String;

  protected
    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
    procedure WriteFooter; override;

  public
    destructor Destroy; override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TBIFFExport
  //----------------------------------------------------------------------------
  // Purpose      : Exports data to Excel in native format.
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  // Limitations  : - Memos and Strings are truncated to 255 chars.
  //----------------------------------------------------------------------------

  TBIFFExport = class(TCustomShellViewExport)
  private
    BIFF_Version: ShortInt; // BIFF Version

    F: File;
    DimensionsOffset: Cardinal;
    RecType, RecLength: Word;

    FTitleRow,
    FTimeStampRow,
    FHeaderRow,
    FFooterRow,
    FColHeadersRow: Integer;

    procedure WriteRecordHeader;
    procedure WriteByteRecord(aRecType: Word; aByte: Byte);
//    procedure WriteWordRecord(aRecType, aWord: Word);
    procedure WriteFlagRecord(aRecType: Word; Flag: Boolean);
    procedure WriteStringRecord(aRecType: Word; S: String);

    procedure WriteBOF;
    procedure WriteEOF;
    procedure WriteDimensions(MaxRows, MaxCols: Integer);

    procedure WriteXF(aParams: array of Byte);
    procedure WriteFont(aFontName: String; aHeight: Integer; aFontStyles: TFontStyles);
    procedure WriteFormat(aFormatString: String);

    procedure WritePrintHeader(S: String);
    procedure WritePrintFooter(S: String);

    procedure WriteColWidth(aColIndex: Byte; NumCharacters: Word);

    procedure WriteData(ColType: TColType; ARow, ACol: Integer; AData: Pointer);
//    procedure WriteBlank;
    procedure WriteNumber;
    procedure WriteLabel(var w: Word; AData: Pointer);
    procedure WriteBoolean;

  protected
    procedure WriteCell(Value: String; Row, Col: Integer; ColType: TColType; ItemInfo: TExportItemInfo);

    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
    procedure WriteFooter; override;

  public
    constructor Create(aOwner: TExportXComponent; aExportFile: String; aExportType: TExportType); override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TCSVExport
  //----------------------------------------------------------------------------
  // Purpose      : Exports data to CSV - Comma Seperated Value format.
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  // Notes        : Although CSV is a universal standard, there are problems
  //                with many program's implementions of them. Furthermore,
  //                in some locales (like Germany) the ListSeparator value is
  //                not a comma but rather a semi-colon. So most-likely some
  //                programs will not be able to import CSV files separated with
  //                semicolons (which is what TExportX will produce in those
  //                Locales, just like Excel does when it exports to CSV in
  //                those Locales)
  //----------------------------------------------------------------------------

  TCSVExport = class(TCustomTextExport)
  private
    function MakeCSV(var Items: TStringList): String;

  protected
    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TTabExport
  //----------------------------------------------------------------------------
  // Purpose      : Exports data to the Tab-Delimited Text Files.
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  //----------------------------------------------------------------------------

  TTabExport = class(TCustomTextExport)
  protected
    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TTextExport
  //----------------------------------------------------------------------------
  // Purpose      : Exports data to a Text File in a Grid format padded with
  //                spaces.
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  // Notes        : If a column is not as wide as the data, the data will get
  //                truncated.
  //----------------------------------------------------------------------------

  TTextExport = class(TCustomTextExport)
  protected
    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
    procedure WriteFooter; override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TDetailedTextExport
  //----------------------------------------------------------------------------
  // Purpose      : Exports data to a Text Files in a Detailed mode.
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  //----------------------------------------------------------------------------

  TDetailedTextExport = class(TCustomTextExport)
  protected
    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
    procedure WriteFooter; override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TDIFExport
  //----------------------------------------------------------------------------
  // Purpose      : Exports data to a DIF (Data Interchange Format) File.
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  //----------------------------------------------------------------------------

  TDIFExport = class(TCustomTextExport)
  private
    FTempDIF: String; // Name of Temp DIF file
    TempF: TextFile;  // Temp File Object
  protected
    procedure WriteTuple(var Items: TStringList);

    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
    procedure WriteFooter; override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TSYLKExport
  //----------------------------------------------------------------------------
  // Purpose      : Exports data to a SYLK (Symbolic Link) File.
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  // Limitations  : - Memos and Strings are truncated to 255 chars.
  //----------------------------------------------------------------------------

  TSYLKExport = class(TCustomTextExport)
  protected
    procedure WriteRow(var Items: TStringList);

    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
    procedure WriteFooter; override;
  end;

  //----------------------------------------------------------------------------
  // Class        : THTMLExport
  //----------------------------------------------------------------------------
  // Purpose      : Exports data to an HTML File.
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  // Uses         : HTML_Template, HTML_CustomTemplate and HTML_Options as well
  //                as regular export properties.
  //----------------------------------------------------------------------------

  THTMLExport = class(TCustomTextExport)
  private
    FSelectedTemplate: String;

    function ConvertToHTML(S: String): String;

    function LinkURLs(S: String): String;
    function EnsureNotEmpty(S: String): String;

    function GetElement(i: Integer): String;

  protected
    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
    procedure WriteFooter; override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TRichTextExport
  //----------------------------------------------------------------------------
  // Purpose      : Exports data to a RichText File (RTF).
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  //----------------------------------------------------------------------------

  TRichTextExport = class(TCustomTextExport)
  private
    function ConvertToRTF(S: String): String;
    procedure WriteColFormatInfo(Cell, Prefix, Suffix, Regular: String);

  protected
    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
    procedure WriteFooter; override;

  public
    constructor Create(aOwner: TExportXComponent; aExportFile: String; aExportType: TExportType); override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TCustomPrintExport
  //----------------------------------------------------------------------------
  // Purpose      : Base Class for Print Export Objects.
  // Functionality: Abstract Class.
  //----------------------------------------------------------------------------

  TCustomPrintExport = class(TCustomExport)
  protected
    PrintText: TextFile;

    procedure WriteHeader(var Columns: TStringList); override;

    procedure HandleError(Exception: TObject); override;

  public
    constructor Create(aOwner: TExportXComponent; aExportFile: String; aExportType: TExportType); override;
    destructor Destroy; override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TPrintExport
  //----------------------------------------------------------------------------
  // Purpose      : Prints data to the printer in a Grid format padded with
  //                spaces using a fixed-width font.
  // Functionality: Called by the TExportXComponent's Print method.
  // Notes        : If a column is not as wide as the data, the data will get
  //                truncated.
  //----------------------------------------------------------------------------

  TPrintExport = class(TCustomPrintExport)
  protected
    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
    procedure WriteFooter; override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TDetailedPrintExport
  //----------------------------------------------------------------------------
  // Purpose      : Prints the data in Detailed mode using a variable pitch font
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  //----------------------------------------------------------------------------

  TDetailedPrintExport = class(TCustomPrintExport)
  protected
    procedure WriteHeader(var Columns: TStringList); override;
    procedure WriteEntry(var Items: TStringList; Entry: Integer); override;
    procedure WriteFooter; override;
  end;

  //----------------------------------------------------------------------------
  // Class        : TExcelExport
  //----------------------------------------------------------------------------
  // Purpose      : Exports data to an Excel File.
  // Functionality: Called by the TExportXComponent's Choose or Execute methods.
  // Limitations  : Memos and Strings are limited to 255 characters.
  //----------------------------------------------------------------------------

// *** Misc. Functions ***

function IsCardinal(S: String): Boolean;
function IsBlankString(S: String): Boolean;

function GetNumDataColumns(ColCount: Integer; aOptions: TExportOptions): Integer;
function GetUnderline(aLength: Integer): String;
function GetSingleColEntry(aEntry: String; aEntryIndex: Integer; aOptions: TExportOptions): String;
function GetToken(S: String; i: Integer): String;

function ExpandTempDir(TheFile: String): String;
function ReinforceSymbol(S: String; Symbol, Reinforcer: Char): String;
function Has(S: String): Boolean;
function ClassExists(aClass: String): Boolean;

function GetExportType(Choice: String): TExportType;

implementation

{$R-}

uses SysUtils, ShellAPI, Forms, Controls, Clipbrd, Printers, CommCtrl,

   // TExportXComponent Units
   EC_Choice;

const
  // ColType Sets
  cNumericColTypes = [ctNumber, ctCurrency];

  // Control Characters
  cTab = Char($09);
  cCR  = Char($0D);
  cLF  = Char($0A);

  // LineFeeds
  CRLF   = #13#10;
  CRLFx2 = CRLF + CRLF;

  // Temp. Constants
  cMaxPathSize = 1024;
  cTempFilePrefix  = 'exp';   // First 3 Characters of Temp File Name

  // Text & Clipboard Constants
  cNumberPostFix = '.'; // This is the character that comes right after the number in a single-column export
  cUnderlineChar = '-';

  // Default ColWidth Spacing
  cDefaultColWidthSpacing     = 1;
  cRTF_DefaultColWidthSpacing = 5;

  // Default Printer Fonts
  cDetailedTitleName = 'Arial';
  cDetailedTitleSize = 16;
  cDetailedTitleStyle: TFontStyles = [fsBold, fsUnderline];

  cDefaultFontName = 'Courier New';
  cDefaultFontSize = 8;
  cDefaultFontStyle: TFontStyles = [];

  cDetailedFontName = 'Courier New';
  cDetailedFontSize = 9;
  cDetailedFontStyle: TFontStyles = [];

  cDetailedFooterName = 'Arial';
  cDetailedFooterSize = 12;
  cDetailedFooterStyle: TFontStyles = [fsItalic];

  cDetailedHeaderName = 'Arial';
  cDetailedHeaderSize = 12;
  cDetailedHeaderStyle: TFontStyles = [fsItalic];

  // SYLK Constants
  cSYLK_FormatIndex_Date     = 0; // SYLK Format Table Index for Date Template
  cSYLK_FormatIndex_Time     = 1; // ... for Time Template
  cSYLK_FormatIndex_DateTime = 2; // ... for DateTime Template
  cSYLK_CurrencyTemplate     = 3; // ... for Currency Template

  // BIFF Misc. Constants
  BIFF_RowNotPresent  = -1;

  // BIFF Font Constants

  BIFF_BaseFontIndex  = 5; // First Font in Font Table we use starts with 5 - due to backwards compatibility with older BIFF versions
  BIFF_DataFont       = 0; // Default Font for Data
  BIFF_ItalicFont     = 1; // Italic Font
  BIFF_DetailedTitle  = 2; // Title Font
  BIFF_ColHeadersFont = 3; // Column Headers Row Font

  // BIFF Format Strings

  BIFF_GeneralFormatString  = 'General';

  // BIFF Format Constants

  BIFF_BaseFormatIndex = $40;
  BIFF_GeneralFormat   = 0;
  BIFF_DateFormat      = 1;
  BIFF_TimeFormat      = 2;
  BIFF_DateTimeFormat  = 3;
  BIFF_CurrencyFormat  = 4;

  BIFF_TitleFormat     = 5;
  BIFF_ItalicFormat    = 6;
  BIFF_ColHeaderFormat = 7;

  BIFF_TimeStampFormat = BIFF_GeneralFormat;
  BIFF_HeaderFormat    = BIFF_GeneralFormat;
  BIFF_FooterFormat    = BIFF_ItalicFormat;

  // BIFF Border and Alignment Constants (all for the same Byte)

  BIFF_CellShaded   = $80; // Bit 7
  BIFF_TopBorder    = $40; // Bit 6
  BIFF_BottomBorder = $20; // Bit 5
  BIFF_RightBorder  = $10; // Bit 4
  BIFF_LeftBorder   = $08; // Bit 3

  BIFF_AlignGeneral = $00; // No Bits  (000b)
  BIFF_AlignLeft    = $01; // Bit 0    (001b)
  BIFF_AlignCenter  = $02; // Bit 1    (010b)
  BIFF_AlignRight   = $03; // Bits 1,0 (011b)
  BIFF_AlignFill    = $04; // Bit 2    (100b)

  // BIFF OpCodes

  BIFF_DIMENSIONS = $0000; // Dimensions
  BIFF_BOF        = $0009; // BOF
  BIFF_EOF        = $000A; // EOF

  BIFF_XF         = $0043; // Known as XF, Extended Format in Later BIFF Versions
  BIFF_FONT       = $0031; // Font Record
  BIFF5_FONT      = $0231; // BIFF5 Font Record
  BIFF_FORMAT     = $001E; // Format Record - for Cell "Picture" Strings

  BIFF_HEADER     = $0014; // Print Header
  BIFF_FOOTER     = $0015; // Print Footer
  BIFF_ROWHEADERS = $002A; // Print Row Headers
  BIFF_GRIDLINES  = $002B; // Print Gridlines

  BIFF_COLWIDTH   = $0024; // ColWidth Record

  BIFF_BLANK      = $0001;
  BIFF_INTEGER    = $0002;
  BIFF_NUMBER     = $0003;
  BIFF_LABEL      = $0004;
  BIFF_BOOLEAN    = $0005;

  // HTML Constants

  cBeginHTML = '<html><head><title>%s</title></head>' + CRLF +
               '<body bgcolor="%s" link="%s" vlink="%s" text="%s" alink="%s">' + CRLF +
               '<font color="%s" face="%s">';
  cEndHTML = '</font></body></html>';

  cHTML_Title        = '<h1>%s</h1>';
  cHTML_LineBreak    = '<br>';
  cHTML_LineBreak_x2 = cHTML_LineBreak + cHTML_LineBreak;
  cHTML_HorzRule     = '<hr>';

  cEmptyCell = '&nbsp;';

  cBeginTable = '<table cellspacing="0" cellpadding="4" border="%d" bgcolor="%s">';
  cEndTable = '</table>';

  cBeginBold   = '<b>';
  cEndBold     = '</b>';

  cBeginColoredRow  = '<tr valign="top" bgcolor="%s">';
  cBeginRow         = '<tr valign="top">';
  cEndRow           = '</tr>';

  cBeginTitleCell = '<td>%s<font color="%s">';
  cEndTitleCell   = '</font>%s</td>';

  cBeginCell = '<td>';
  cEndCell   = '</td>';

  // HTML Templates
  // Template Format = BGColor;LinkColor;VLinkColor;ALinkColor;TextColor;TextFont;ColHeadersBGColor;ColHeadersFontColor;TableBGColor;TableFontColor;OddRowColor

  cHTML_Templates: THTML_Templates = (
  {Classic } '#333399;#69EF7D;#FF00FF;#00FF00;#FFFFFF;Arial,Helvetica;#FF0000;#FFFFFF;##006BCE;#FFFFFF;007AEC',
  {Simple  } '#FFFFFF;blue;purple;red;#004080;Arial,Helvetica;#336699;#FFFFFF;#FFFFFF;#000000;#FFFFCF',
  {Plain   } '#FFFFFF;blue;purple;red;#000000;;#068FE6;#FFFFFF;#FFFFFF;#000000;#FFFCD9',
  {Murky   } '#317676;#FFE760;#CCCC39;#FFFF00;#FFFFFF;Arial,Helvetica;#3A9393;#FFFFFF;#004040;#FFE760;#000000',
  {MSMoney } '#FFFFFF;blue;purple;red;#000080;Arial,Helvetica;#CEC6B5;#000000;#DEE7DE;#000000;FFFBF0',
  {Colorful} '#339966;#0066CC;#923FC1;red;#FFFFFF;Geneva,Arial,Helvetica;#CC0033;#FFFFFF;#FFFFFF;#000000;#FAF100',
  {Gray    } '#FFFFFF;blue;purple;red;#000000;Arial,Helvetica;#808080;#FFFFFF;#FFFFFF;#000000;#EEEEEE',
  {Olive   } '#FFFFFF;blue;purple;red;#000000;Verdana,Arial,Helvetica;#CFC890;#000000;#FFFFFF;#5F605F;#FFFFCF',
  {BW      } '#FFFFFF;blue;purple;red;#000000;Arial,Helvetica;#000000;#FFFFFF;#FFFFFF;#000000;#F3F3F3');

  cHTML_NumTokensRequired   = 11; // Number of Tokens Required
  cHTML_NumSemicolons       = cHTML_NumTokensRequired - 1; // Number of Semicolons required in a valid Template

  cHTML_BGColor             = 1; // Background Colors
  cHTML_Link                = 2; // Default Link Colors
  cHTML_VLink               = 3; // Visited Link Colors
  cHTML_ALink               = 4; // Active Link Colors
  cHTML_Text                = 5; // Default Text Colors (used for all text outside of the table)
  cHTML_TextFont            = 6; // The Text Font (used for all text)
  cHTML_ColHeadersBGColor   = 7; // The Column Headers Row Background Color
  cHTML_ColHeadersFontColor = 8;  // The Column Headers Row Font Color
  cHTML_TableBGColor        = 9;  // The Table Background Color (used for all of the table except the column headers row)
  cHTML_TableFontColor      = 10; // The Table Font Color (used for all text in table except the column headers row)
  cHTML_OddRowBGColor       = 11; // Odd Row BGColor (used for odd rows in table, if it should be the same color as the even rows, leave this value blank ie. ';')

  // RichText Constants
  cRTF_LF   = '\par';
  cRTF_LFx2 = cRTF_LF + cRTF_LF;

  cRTF_Header  =

'{\rtf1\ansi\ansicpg1252\uc1\deff0\deflang1033\deflangfe1033{\fonttbl{\f0\froman\fcharset0\fprq2{' +
'\*\panose 02020603050405020304}Times New Roman;}{\f1\fswiss\fcharset0\fprq2{\*\panose 020b0604020202020204}Arial;}}{\colortbl;\red0\green0\blue0;' + CRLF +

'\red0\green0\blue255;\red0\green255\blue255;\red0\green255\blue0;\red255\green0\blue255;\red255\' +
'green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;' + CRLF +

'\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;\red192\green192\blue192;}{' +
'\stylesheet{\nowidctlpar\widctlpar\adjustright\fs20\cgrid\snext0 Normal;}{\s1\keepn\nowidctlpar\widctlpar\outlinelevel0\adjustright\b\f1\fs28\cgrid' + CRLF +

'\sbasedon0\snext0 heading 1;}{\*\cs10\additive Default Paragraph Font;}}{\info{\nofcharsws0}{\vern71}}' +
'\widowctrl\ftnbj\aenddoc\hyphcaps0\formshade\viewkind4\viewscale109\viewzk2\pgbrdrhead\pgbrdrfoot\fet0\sectd\linex0\endnhere\sectdefaultcl ' + CRLF +

'\margl%d\margr%d\margt%d\margb%d' + CRLF +

'{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang' + CRLF +

'{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1' +
'\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}{\*\pnseclvl5\pndec\pnstart1\pnindent720\pnhang{\pntxtb (}' + CRLF +

'{\pntxta )}}{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm' +
'\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl9' + CRLF +

'\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}\pard';

cRTF_Title = '\plain\s1\keepn\nowidctlpar\widctlpar\outlinelevel0\adjustright\b\f1\fs28\cgrid {%s}\fs16' + cRTF_LFx2;

cRTF_BeginText = '\pard\plain\nowidctlpar\widctlpar\adjustright\fs20\cgrid {\f1\fs16';

cRTF_TimeStamp = '%s' + cRTF_LFx2;
cRTF_UserHeader = '%s' + cRTF_LF;

cRTF_ColFormatHeadingCell   = '\brdrs\brdrw30\brdrcf1\clbrdrb\brdrs\brdrw15\brdrcf1\cltxlrtb\cellx%d';
cRTF_ColFormatHeadingPrefix = '}\trowd\trgaph108\trleft-108\trbrdrt\brdrs\brdrw30\brdrcf1\trbrdrb\brdrs\brdrw30\brdrcf1\clvertalt\clbrdrt';
cRTF_ColFormatHeadingSuffix = '\pard\nowidctlpar\widctlpar\intbl\adjustright{\b\f1\fs16';
cRTF_ColFormatHeading       = cRTF_ColFormatHeadingCell + '\clvertalt\clbrdrt';

cRTF_AfterCell = '\cell ';

cRTF_ColFormatInitialRowCell   = '\clvertalt\cltxlrtb\cellx%d';
cRTF_ColFormatInitialRowPrefix = '}\pard\nowidctlpar\widctlpar\intbl\adjustright {\i\f1\fs16\row}\trowd\trgaph108\trleft-108\trbrdrt\brdrs\brdrw30\brdrcf1\trbrdrb\brdrs\brdrw30\brdrcf1';
cRTF_ColFormatInitialRowSuffix = '\pard\nowidctlpar\widctlpar\intbl\adjustright {\f1\fs16';

cRTF_BeforeRow = '}\pard\nowidctlpar\widctlpar\intbl\adjustright{\f1\fs16\row}\pard\nowidctlpar\widctlpar\intbl\adjustright{\f1\fs16';

cRTF_ColFormatFinalRowCell   = '\brdrs\brdrw30\brdrcf1\cltxlrtb\cellx%d';
cRTF_ColFormatFinalRowPrefix = '}\pard\nowidctlpar\widctlpar\intbl\adjustright{\f1\fs16\row}\trowd\trgaph108\trleft-108\trbrdrt\brdrs\brdrw30\brdrcf1\trbrdrb\brdrs\brdrw30\brdrcf1\clvertalt\clbrdrb';
cRTF_ColFormatFinalRowSuffix = '\pard\nowidctlpar\widctlpar\intbl\adjustright{\f1\fs16';
cRTF_ColFormatFinalRow       = cRTF_ColFormatFinalRowCell + '\clvertalt\clbrdrb';

cRTF_Footer = '}\pard\nowidctlpar\widctlpar\intbl\adjustright{\f1\fs16\row}\pard\nowidctlpar\widctlpar\adjustright{\f1\fs16\par %s}}';

var
  TempFiles: TStringList;
  SearchRec: TSearchRec;
  FUserTempDir: String;

// *** TAutoCleanupStringList Methods ***

procedure TAutoCleanupStringList.Clear;
var
  i: Integer;
begin
  for i := 0 to Count - 1 do Objects[i].Free; // Free All Objects in StringList

  inherited; // Clear StringList
end;

procedure TAutoCleanupStringList.AutoFree;
var
  i: Integer;
begin
  for i := 0 to Count - 1 do Objects[i].Free; // Free All Objects in StringList
  Free; // Free StringList
end;

// *** TMyReg Methods ***
function TMyReg.OK_ReadOnly(const Key: String): Boolean;
  function IsRelative(const Value: string): Boolean;
  begin
    Result := not ((Value <> '') and (Value[1] = '\'));
  end;
var
  TempKey: HKey;
  S: string;
  Relative: Boolean;
begin
  S := Key;
  Relative := IsRelative(S);

  if not Relative then Delete(S, 1, 1);
  TempKey := 0;
  Result := RegOpenKeyEx(GetBaseKey(Relative), PChar(S), 0,
      KEY_READ, TempKey) = ERROR_SUCCESS;
  if Result then
  begin
    if (CurrentKey <> 0) and Relative then S := CurrentPath + '\' + S;
    ChangeKey(TempKey, S);
  end;
end;

// *** TPrint_Options Methods ***

constructor TPrint_Options.Create;
  procedure InitFont(var aFont: TFont; aFontName: String; aFontSize: Integer; aFontStyle: TFontStyles);
  begin
    aFont := TFont.Create;
    with aFont do
    begin
      Name  := aFontName;
      Size  := aFontSize;
      Style := aFontStyle;
    end;
  end;

begin
  // Initialize
  InitFont(FDetailedTitle, cDetailedTitleName, cDetailedTitleSize, cDetailedTitleStyle);
  InitFont(FDefaultFont, cDefaultFontName, cDefaultFontSize, cDefaultFontStyle);
  InitFont(FDetailedHeader, cDetailedHeaderName, cDetailedHeaderSize, cDetailedHeaderStyle);
  InitFont(FDetailedFooter, cDetailedFooterName, cDetailedFooterSize, cDetailedFooterStyle);
  InitFont(FDetailedFont, cDetailedFontName, cDetailedFontSize, cDetailedFontStyle);
end;

destructor TPrint_Options.Destroy;
begin
  // Free Font Objects
  FDetailedTitle.Free;
  FDefaultFont.Free;
  FDetailedHeader.Free;
  FDetailedFooter.Free;
  FDetailedFont.Free;

  inherited;
end;

// *** TXL_Options Methods ***

constructor TXL_Options.Create;
begin
  // Initialize
  FPageFooter := cExcelRightFooter; // Get default value from EC_Strings.pas
end;

// *** TRTF_Options Methods ***

constructor TRTF_Options.Create;
begin
  // Initialize
  FLeftMargin   := cRTF_DefaultLeftRightMargin;
  FRightMargin  := cRTF_DefaultLeftRightMargin;

  FTopMargin    := cRTF_DefaultTopBottomMargin;
  FBottomMargin := cRTF_DefaultTopBottomMargin;
end;

// *** TSharewareExportComponent Methods ***

constructor TSharewareExportComponent.Create(AOwner: TComponent);
begin
{$IFDEF NotRegistered} // Quit if UnRegistered Version and Not Running Delphi
  if FindWindow('TAppBuilder', nil) <= 0 then
  begin
    MessageDlg(Format(cIDENotRunningError, [ClassName]), mtError, [mbOk], 0);
    Application.Terminate;
  end;
{$ENDIF}

  inherited;
end;

// *** TExportXComponent Methods ***

constructor TExportXComponent.Create(AOwner: TComponent);
var
  i: Integer;
begin
  inherited;

  // Initialize
  FAppMinimized := False;

  // Set Property Defaults
  FExportType     := xMicrosoft_Excel;       // Set Excel as Default Export Type
  FHTML_Template  := htSimple;
  FViewOnly       := True;
  FOptions        := cDefOptions;            // Set Default Options
  FHTML_Options   := cDefHTML_Options;       // Set Default HMTL Options
  FTruncateSymbol := cDefaultTruncateSymbol; // Set Default Truncate Symbol

  // Create Objects
  FCaptions          := TStringList.Create;
  FColWidths         := TStringList.Create;
  FHeader            := TStringList.Create;
  FFooter            := TStringList.Create;
  FFormatFooter      := TStringList.Create;
  FExportColumns     := TStringList.Create;
  FPrint_Options     := TPrint_Options.Create;
  FXL_Options        := TXL_Options.Create;
  FRTF_Options       := TRTF_Options.Create;

  // Allow All Export Types by Default
  for i := Ord(Low(TExportType)) to Ord(High(TExportType)) do
    Include(FAllowedTypes, TExportType(i));

  // Hook Application's OnActivate Event
  FOldActivate := Application.OnActivate;
  Application.OnActivate := OnActivate;
end;

procedure TExportXComponent.Loaded;
begin
  inherited;

  // Initialize FLastExportType
  FLastExportType := FExportType; // Set Last Export Type to default export type
end;

procedure TExportXComponent.OnActivate(Sender: TObject);
begin
  // First Call Old Activate Event
  if Assigned(FOldActivate) then
    FOldActivate(Sender);

  // Restore App that was minmized by TCustomExport.Destroy (if any)    
  RestoreApp;
end;

function TExportXComponent.Strip(S: String): String;
var
  i: Integer;
begin
  if Length(S) > 0 then
  begin
    i := 1;
    repeat
      if Ord(S[i]) < 32 then
        Delete(S, i, 1) // Remove control character
      else Inc(i)
    until i > Length(S);
  end;

  Result := S;
end;

procedure TExportXComponent.TrimTrailingBlankStrings(var S: TStrings);
var
  i: Integer;
begin
  // Trim Trailing Blank Strings
  with S do
    for i := Count - 1 downto 0 do
      if IsBlankString(Strings[i]) then
        Delete(i)
      else Break;
end;

function TExportXComponent.GetDefaultColWidth(ColIndex: Integer): Integer;
begin
  // Get Length of Column
  if NumberRows in FOptions then
    Result := Length(FExportColumns[ColIndex + 1])
  else Result := Length(FExportColumns[ColIndex]);

  // Add Spacing
  Result := Result + FDefaultColWidthSpacing;
end;

destructor TExportXComponent.Destroy;
begin
  // Unhook OnActivate Event
  Application.OnActivate := FOldActivate;

  // Free Objects
  FCaptions.Free;
  FColWidths.Free;
  FHeader.Free;
  FFooter.Free;
  FFormatFooter.Free;
  FExportColumns.Free;

  FPrint_Options.Free;
  FXL_Options.Free;
  FRTF_Options.Free;

  inherited;
end;

// *** procedures which must be overriden by descendants ***

procedure TExportXComponent.InitExport;
begin
//
end;

procedure TExportXComponent.CleanUpExport;
var
  i: Integer;
begin
  // Free All Objects in FExportColumns List
  with FExportColumns do
    for i := 0 to Count - 1 do
      if Assigned(Objects[i]) then
        Objects[i].Free;
end;

function TExportXComponent.GetColumnCaption(aIndex: Integer): String;
begin
  if aIndex < FCaptions.Count then     // If there's a user-defined caption,
    Result := FCaptions[aIndex]        // use it,
  else Result := GetColValue(aIndex) ; // Otherwise, use the default
end;

function TExportXComponent.GetExportColumns(var aColumns: TStringList): TStringList;
begin
  // Common Stuff
  Result := aColumns;       // Return Columns
  aColumns.Clear;           // Clear Old Columns Values
end;

procedure TExportXComponent.InitExportItems;
begin
  // Initialize
  FCount.CurrentRow   := 0;
  FCount.RowsExported := 0;
end;

function TExportXComponent.CurrentItemSelected: Boolean;
begin
  Result := True; // Default to True (some export components cannot select items so "every item is selected"
end;

procedure TExportXComponent.SkipToNextItem;
begin
  Inc(FCount.CurrentRow);
end;

procedure TExportXComponent.GetNextExportItem(var aExportItem: TStringList);
begin
  // If SelectedRowsOnly Set, Skip over items we shouldn't export
  if not ((FCount.SelectedItems = 0) and (ExportAllWhenNoneSelected in FOptions)) then // unless there's nothing selected and ExportAllWhenNoneSelected is Set
    while (SelectedRowsOnly in FOptions) and not CurrentItemSelected do
      SkipToNextItem;

  // Clear Contents of StringList
  aExportItem.Clear;

  // Get Export Item Data
  GetExportItemData(aExportItem);

  Inc(FCount.RowsExported); // Increment Number of Rows Exported
  Inc(FCount.CurrentRow);   // Increment Current Row

  // If Numbering Rows, Insert Line Number
  if NumberRows in FOptions then
    aExportItem.Insert(0, IntToStr(FCount.RowsExported));
end;

function TExportXComponent.MoreExportItems: Boolean;
begin
  with FCount do
    Result := RowsExported < ExportItems;
end;

function TExportXComponent.HasExportableColumns: Boolean;
var
  S: TAutoCleanupStringList;
  i: Integer;
begin
  Result := False; // Assume Failure

  S := TAutoCleanupStringList.Create;
  try
    GetExportColumns(TStringList(S));
    if S.Count > 0 then                 // Preliminary Check: Are there any columns?
      for i := 0 to S.Count - 1 do      // Detailed Check: Are any of them...
        if not IsBlankString(S[i]) then // ...not blank?
          Result := True;               // If So, we've got exportable columns!

  finally
    S.AutoFree; // Free Stringlist and associated Objects
  end;

  // Display Error Message if no columns to export
  if not Result then
    MessageDlg(cNoDataError + cNoColumnsError, mtWarning, [mbOk], 0);
end;

// ***

function TExportXComponent.ExportAppInstalled(ExportType: TExportType): Boolean;
  function Associated(aExt: String): Boolean;
  { This is a bit tricky:
      1) Find out if the extension is listed in HKEY_CLASSES_ROOT.
      2) Get it's class name.
      3) See if class has a shell key
      4) if class has "open" key then return True
  }
  const
    cShellKey = 'shell';
    cOpenKey  = 'open';
    cQuickTimeDIF = 'QuickTime.dif';
  var
    ViewerClass: String;
  begin
    Result := False; // Assume Failure

    with TMyReg.Create do
    try
      RootKey := HKEY_CLASSES_ROOT;

      if OK_ReadOnly('.' + aExt) then // Attempt to Open Extension Association key
      begin
        ViewerClass := ReadString(''); // Get Class Name

        // If QuickTime took over the .DIF extension, we report export app not installed for DIF files
        if (ExportType = xDIF) and (CompareText(ViewerClass, cQuickTimeDIF) = 0) then
        begin
          Result := False;
          Exit;
        end;

        CloseKey; // Close Extension Association Key

        // Get Extension's Associated Class "Open" Viewer
        Result := OK_ReadOnly(ViewerClass) and
                  OK_ReadOnly(cShellKey) and
                  OK_ReadOnly(cOpenKey);
      end;

    finally
      CloseKey;
      Free;
    end;
  end;

begin
  try
    // If we have an extension in the ExportType Extensions List, then see if it's associated
    if Length(cExportTypeExtensions[ExportType]) > 0 then
      Result := Associated(cExportTypeExtensions[ExportType])
    else Result := True; // All other Apps are installed by default

  except
    Result := False; // Fail on exception
  end;
end;

procedure TExportXComponent.PopulateStrings(S: TStrings);
var
  i: Integer;
begin
  S.Clear;
  for i := Integer(Low(TExportType)) to Integer(High(TExportType)) do
    if TExportType(i) in FAllowedTypes then // As long as this type is allowed,
      S.Add(cExportTypes[TExportType(i)]); // Add Export Type Description to Strings
end;

procedure TExportXComponent.SetTitle(Value: String);
begin
  FTitle := Strip(Value); // Make sure there is no nasty control characters in it
end;

procedure TExportXComponent.SetHeader(Value: TStrings);
begin
  TrimTrailingBlankStrings(Value);
  FHeader.Assign(Value); // Assign Value
end;

procedure TExportXComponent.SetFooter(Value: TStrings);
begin
  TrimTrailingBlankStrings(Value);
  FFooter.Assign(Value); // Assign Value
end;

procedure TExportXComponent.SetTruncateSymbol(Value: String);
begin
  FTruncateSymbol := Trim(Value);
end;

procedure TExportXComponent.SetCaptions(Value: TStrings);
begin
  TrimTrailingBlankStrings(Value);
  FCaptions.Assign(Value); // Assign Value
end;

procedure TExportXComponent.SetColWidths(Value: TStrings);
const
  cInvalidColWidthError = 'ColWidths Property Error' + CRLF + CRLF +
                          '"%s" is an Invalid Column Width. It must be a ' +
                          'Positive Integer Number greater than Zero.';
var
  i: Integer;
begin
  // Remove all trailing blank lines
  if Value.Count > 0 then
    for i := Value.Count - 1 downto 0 do
      if IsBlankString(Value[i]) then
        Value.Delete(i)
      else break;

  // First check to see that all values are cardinal numbers or blank lines
  for i := 0 to Value.Count - 1 do
  begin
    Value[i] := Trim(Value[i]); // Trim it

    if not (IsBlankString(Value[i]) or                                // If It's Blank
            (IsCardinal(Value[i]) and (StrToInt(Value[i]) > 0))) then // or not Integer > 0
      raise Exception.CreateFmt(cInvalidColWidthError, [Value[i]]);
  end;

  // Assign Value
  FColWidths.Assign(Value);
end;

procedure TExportXComponent.SetPrint_Options(Value: TPrint_Options);
begin
  FPrint_Options.Assign(Value);
end;

procedure TExportXComponent.SetXL_Options(Value: TXL_Options);
begin
  FXL_Options.Assign(Value);
end;

procedure TExportXComponent.SetRTF_Options(Value: TRTF_Options);
begin
  FRTF_Options.Assign(Value);
end;

procedure TExportXComponent.SetOptions(Value: TExportOptions);
begin
  if FOptions <> Value then
    FOptions := Value - FIllegalOptions; // Ensure no illegal options are selected
end;

procedure TExportXComponent.SetHTML_Options(Value: THTML_Options);
begin
  if FHTML_Options <> Value then
    FHTML_Options := Value;
end;

procedure TExportXComponent.SetHTML_CustomTemplate(Value: String);
const
  cInvalidTemplateError = 'Invalid Template! Please correct it.' + CRLF + CRLF +
                          'You should have "%d" semicolons but instead you have "%d"';
var
  i, n: Integer;
begin
  // If Blank, then don't worry about counting tokens
  if IsBlankString(Value) then
  begin
    FHTML_CustomTemplate := '';
    Exit;
  end;

  // Count # of Tokens
  n := 0;
  for i := 1 to Length(Value) do
    if Value[i] = ';' then
      Inc(n);

  if n = cHTML_NumSemiColons then  // If We have correct number of tokens,
    FHTML_CustomTemplate := Value  // then set the property,
  else raise Exception.CreateFmt(cInvalidTemplateError, [cHTML_NumSemiColons, n]); // else raise exception
end;

function TExportXComponent.Execute: Boolean;
var
  E: TCustomExport;
  ViewOnScreen: Boolean;
  OutputFile: String;
begin
  Result := False; // Assume failure

  try
    if HasData then // If Exportable (ie. it has data)
    try
      Application.ProcessMessages;

      // Set ViewOnScreen property (not always the same as ViewOnly)
      if FExportType = xClipboard then
        ViewOnScreen := True // Clipboard never exports to file
      else
        if not ExportAppInstalled(FExportType) then // Always export to file if no App installed
        begin
            ViewOnScreen := False;             // then prompt user for filename
            FExportFile  := '';
        end
        else ViewOnScreen := FViewOnly;      // otherwise, it's up the ViewOnly Properly

      // If not ViewOnScreen and no export file specified, prompt user for Export File Name
      if (not ViewOnScreen) and IsBlankString(FExportFile) then
        with TSaveDialog.Create(nil) do
        try
          DefaultExt := cExportTypeExtensions[FExportType]; // Get Extension
          Options    := cFileDialogOptions;
          Title      := Format(cExportTitle, [cExportTypes[FExportType]]);
          Filter     := Format(cFilesFilter, [cExportTypes[FExportType],
                                              DefaultExt, DefaultExt]);

          if Execute then           // If User Gave a File Name,
            FExportFile := FileName // use that one,
          else Exit;                // otherwise, he pressed cancel - so cancel Execute

        finally
          Free;
        end;

      // Handle View Only Flag
      if ViewOnScreen then
        OutputFile := '' // View Only
      else OutputFile := FExportFile;

      // Get Number of Export Items
      FCount.ExportItems := GetNumExportableItems;

      // Create the Export Object
      E := nil;                      // Initialize It
      case FExportType of
        xHTML                 : E := THTMLExport.Create(Self, OutputFile, xHTML);
        xMicrosoft_Excel      : E := TBIFFExport.Create(Self, OutputFile, xMicrosoft_Excel);
        xMicrosoft_Word       : E := TRichTextExport.Create(Self, OutputFile, xMicrosoft_Word);
        xText_Comma_Delimited : E := TCSVExport.Create(Self, OutputFile, xText_Comma_Delimited);
        xText_Tab_Delimited   : E := TTabExport.Create(Self, OutputFile, xText_Tab_Delimited);
        xDIF                  : E := TDIFExport.Create(Self, OutputFile, xDIF);
        xSYLK                 : E := TSYLKExport.Create(Self, OutputFile, xSYLK);
        xRichText             : E := TRichTextExport.Create(Self, OutputFile, xRichText);

        // Export Types that have a Detailed Mode...
        xText:
          if DetailedMode in FOptions then
            E := TDetailedTextExport.Create(Self, OutputFile, xText)
          else E := TTextExport.Create(Self, OutputFile, xText);

        xClipboard:
          if DetailedMode in FOptions then
            E := TDetailedClipboardExport.Create(Self, OutputFile, xClipboard)
          else E := TClipboardExport.Create(Self, OutputFile, xClipboard);
      end;

      // Export the Associated List View
      with E do
      try
        // Trigger Begin Export Event
        if Assigned(FOnBeginExport) then
          FOnBeginExport(Self);

        Result := DoExport; // Export the List View

      finally
        Free; // Free the Export Object
      end;

      // Trigger Export Finished Event on Success
      if Result and Assigned(FOnExportFinished) then
          FOnExportFinished(Self);

    finally
      FExportFile := '';

    end;

  finally
    // Trigger OnExportFailed
    if not Result and Assigned(FOnExportFailed) then
      FOnExportFailed(Self);
  end;
end;

function TExportXComponent.Print: Boolean;
var
  PrintObject: TCustomPrintExport;
begin
  // Initialize
  Result := False;
  Application.ProcessMessages;

  if HasData then
  begin
    // Get Number of Export Items
    FCount.ExportItems := GetNumExportableItems;

    // Choose Print Format
    if DetailedMode in FOptions then
      PrintObject := TDetailedPrintExport.Create(Self, '', xHTML) // xHTML is just a placeholder
    else PrintObject := TPrintExport.Create(Self, '', xHTML);     // xHTML is just a placeholder

    // Print
    with PrintObject do
    try
      Result := DoExport; // Export the List View (returns false on failure)
    finally
      Free; // Free the Export Object
    end;
  end;

  // Trigger OnPrintFailed (if Applicable)
  if not Result and Assigned(FOnPrintFailed) then
    FOnPrintFailed(Self);
end;

function TExportXComponent.Choose: Boolean;
begin
  Result := False; // assume failure

  if HasData then // If an Exportable Object (ie. Object has data)
    with TExportChoiceForm.Create(Self) do
    try
      SelectedExportType := FLastExportType; // Set to Last-Used Export Type
      NumExportItems     := GetNumExportableItems;
      PopulateStrings(ExportTypeComboBox.Items); // Populate Export Choices Combobox

      Result := ShowModal = mrOk;                 // Show Form;

      if Result then // If User pressed Ok,
      begin
        FViewOnly   := ViewOnlyChoosen; // Set View Only Property
        FExportFile := ''; // Make Sure Export File is blank, so we can dialog the user
        FLastExportType := SelectedExportType; // Set Last Export Type
        FExportType     := SelectedExportType; // Set Export Type
        Result := Execute; // Export Now
      end;

    finally
      Free;
    end;

  // Trigger OnExportFailed Event (if Applicable)
  if not Result and Assigned(FOnExportFailed) then
    FOnExportFailed(Self);
end;

procedure TExportXComponent.RestoreApp;
begin
  if FAppMinimized then
  begin
    Application.Restore;
    FAppMinimized := False;
  end;
end;

procedure TExportXComponent.SetTempDir(TempDir: String);
begin
  FUserTempDir := TempDir;
end;

// *** TCustomExport Methods ***

constructor TCustomExport.Create(aOwner: TExportXComponent; aExportFile: String; aExportType: TExportType);
begin
  // Assume Failure
  FCreatedSuccessfully := False;

  // Abort if no Data
  if not aOwner.HasData then Abort;

  // Initialize
  FUserCancelled          := False;
  FNumEntriesExported     := 0;

  // Create the FColumns Object
  FColumns := TStringList.Create;

  // Store the Construction Parameters
  FOwner        := AOwner;        // Store Owner
  FExportFile   := AExportFile;   // Store File name for Descendant's methods to access
  FExportType   := AExportType;   // Store Export Type

  // Initialize FOwner Properties
  FOwner.FDefaultColWidthSpacing := cDefaultColWidthSpacing;

  // Set Export Type Extension
  FExtension := cExportTypeExtensions[aExportType];

  // Flag Export To File if a File was specified
  FExportToFile := Has(aExportFile);

  // Show Progress Form (if Progress)
  if ShowProgress in FOwner.Options then
    FProgForm := CreateProgress(cExportingCaption, '', cChooseDialogStatusMsg, aOwner.FCount.ExportItems);

  // Flag Export Object Created Successfully
  FCreatedSuccessfully := True;
end;

destructor TCustomExport.Destroy;
var
  fwHnd: HWnd;
begin
  if FCreatedSuccessfully then
  begin
    // Free MaxColWidths Array
    FreeMem(PMaxColWidths, FColumns.Count * SizeOf(Integer));

    // Free Columns StringList
    FColumns.Free;

    // Close & Free Progress Form
    if Assigned(FProgForm) then
      FProgForm.Free;

    // Bring Viewer Application to Foreground (if any)
    if not FExportToFile and (FViewerHandle <> 0) then
    begin
      Sleep(250); // Try to Avoid the "viewer app not in foreground after progress form closes" problem
      SetForegroundWindow(FViewerHandle);
      ShowWindow(FViewerHandle, SW_MINIMIZE);

      // Minimize Application (if it's obstructing the view of the exported document)
      fwHnd := GetForegroundWindow;                    // Get Foreground Window
      if  ((fwHnd = Application.Handle) or             // If it's the Application Handle
          ((fwHnd = TForm(FOwner.Owner).Handle))) then // or the component's owner form then
      begin
        Application.Minimize;                          // Minimize It
        FOwner.FAppMinimized := True;                  // Flag Minimized
      end;
    end;
  end;

  inherited;
end;

function TCustomExport.Pad(S: String; aLength: Integer; RightAlignNumbers: Boolean): String;
// If aLength > Length(S) then it pads S with spaces.
// If aLength < Length(S) then it truncates S.
const
  cSpace = ' '; // Pad Character
var
  i, j: Integer;
begin
  // Remove Control Characters
  S := FOwner.Strip(S);

  // Truncate or Allocate Enough Space for Padding
  with FOwner do
    if (Length(S) > aLength) and                // If there's not enough room and
       (Length(FTruncateSymbol) < aLength) then // there's enough room for the the truncate symbol,
      Result := Copy(S, 1, aLength - Length(FTruncateSymbol)) + FTruncateSymbol // Truncate using Truncate Symbol (if any)
    else
    begin
      Result := S;
      SetLength(Result, aLength); // Otherwise allocate space for padding or truncate w/o truncate symbol
    end;

  // Pad (if necessary)
  if RightAlignNumbers and (aLength > Length(S)) and IsNumber(S) then // If Number then,
  begin                                         // right-align
    j := aLength - Length(S) + 1; // Calculate Start of Number Pos

    // Pad Left-Side with spaces
    for i := 1 to j do
      Result[i] := cSpace;

    for i := j to aLength do
      Result[i] := S[i - j + 1];
  end
  else // Otherwise, Left-Align if not number
    for i := Length(S) + 1 to aLength do
      Result[i] := cSpace;
end;

// Padded Export Methods
function TCustomExport.GetPaddedHeader(var Columns: TStringList): String;
var
  S: String;
  i: Integer;
begin
  inherited;

  Result := '';

  with FOwner do
  begin
    // Write Title
    if Has(FTitle) then
    begin
      Result := FTitle + CRLF;
      Result := Result + GetUnderline(Length(FTitle)) + CRLF;
      Result := Result + CRLF;
    end;

    // Write TimeStamp
    if TimeStamp in FOptions then
      Result := Result + GetTimeStamp + CRLFx2;

    // Write Header (and strip control characters excep Tab, CRLFs)
    if Has(FHeader.Text) then
      Result := Result + StripConvert(FHeader.Text, cTab, CRLF) + CRLF;

    // Write Column Header
    if ShowColHeaders in FOptions then // If we're not hiding col headers
    begin
      S := '';
      for i := 0 to Columns.Count - 1 do
        S := S + Pad(Columns[i], GetColWidth(i), True) + ' ';

      SetLength(S, Length(S) - 1); // Truncate the final space
      Result := Result + S + CRLF;
      FUnderlineLength := Length(S); // Get Length of Underline
      Result := Result + GetUnderline(FUnderlineLength) + CRLF; // Write Underline
    end;
  end;
end;

function TCustomExport.GetPaddedEntry(var Items: TStringList; Entry: Integer): String;
var
  S: String;
  i: Integer;
begin
  inherited;

  // Write Item
  Result := '';
  S := '';
  for i := 0 to Items.Count - 1 do
    S := S + Pad(GetDisplayString(Items, i), GetColWidth(i), True) + ' ';

  SetLength(S, Length(S) - 1); // Truncate the final space
  Result := Result + S + CRLF;
end;

function TCustomExport.GetPaddedFooter: String;
begin
  inherited;

  // Write Footer
  Result := '';
  with FOwner do
    if Has(FFormatFooter.Text) then
    begin
      // If we Hide ColHeaders, then don't draw the '-------' separator either
      if ShowColHeaders in FOptions then
        Result := Result + GetUnderline(FUnderlineLength) + CRLFx2; // Write Underline

      // Write Footer (and strip control characters excep Tab, CRLFs)
      Result := Result + StripConvert(FFormatFooter.Text, cTab, CRLF) + CRLF;
    end;
end;

function TCustomExport.StripConvert(S, sTab, sNewLine: String): String;
// This function Strips control characters and converts Tabs and CRLFs to special
// RTF symbols
const
  cControlChars = [Char(0)..Char(31)];
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(S) do
    if S[i] in cControlChars then
      case S[i] of
        cTab : Result := Result + sTab;
        cLF  : Result := Result + sNewLine;
      end
    else Result := Result + S[i]; // If not a control character, add it to the result string
end;

function TCustomExport.StripInvalidChars(S, ValidChars: String): String;
var
  i: Integer;
begin
  // Return String Stripped of Invalid Characters
  Result := '';
  for i := 1 to Length(S) do           // For Each Character in the String,
    if Pos(S[i], ValidChars) > 0 then  // If it is in the Valid Characters String,
      Result := Result + S[i];         // Copy it to the Result String
end;

function TCustomExport.StripNumberFormat(S, RawNumber: String): String;
// -----------------------------------------------------------------------------
// Removes the formatting info from a number. If S is in Scientific Notation,
// the RawNumber is returned.
// -----------------------------------------------------------------------------
// S        : The Formatted String.
// RawNumber: Raw Number String to return if Stripping does not produce a valid
//            number string. 
const
  cNumbers = '0123456789';
var
  S1, ValidChars: String;
  i: Integer;
begin
  Result := RawNumber;        // Assume Failure
  S := Trim(S);               // Trim
  if Length(S) = 0 then Exit; // Quit if nothing there

  // Strip ThousandSeparators
  S1 := '';
  for i := 1 to Length(S) do
    if S[i] <> ThousandSeparator then
      S1 := S1 + S[i];

  // Strip non-valid number characters
  ValidChars := cNumbers + DecimalSeparator;    // Construct ValidChars String
  if Pos(S1[1], '-' + ValidChars) = 0 then Exit; // 1st Character must be a ValidChar or '-' Sign

  // Remaining Characters must be ValidChars
  for i := 2 to Length(S1) do
   if Pos(S1[i], ValidChars) = 0 then Exit;

  // If we didn't abort yet, we've got a valid number string to return
  Result := S1;
end;

procedure TCustomExport.CalculateMaxColumnLength;
var
  i: Integer;
begin
  // Calculate Length of Longest Column Header and store it in FMaxColumnLength
  FMaxColumnLength := 0;
  with FColumns do
    for i := 0 to Count - 1 do
      if Length(Strings[i]) > FMaxColumnLength then
        FMaxColumnLength := Length(Strings[i]);
end;

function TCustomExport.GetRoundedNumber(RawNumber: String; ItemInfo: TExportItemInfo): String;
// Returns either the Raw Number or if there is a display string, it tries to
// figure strip it and return the stripped number. The difference is that the
// stripped number may have been rounded to a specified number of digits after
// the decimal place - or well anything else might have happened to the number
// but we strip everything so it looks like a Raw Number, even though it might
// not match the original raw number.
begin
  if Assigned(ItemInfo) then
    Result := StripNumberFormat(ItemInfo.DisplayString, RawNumber)
  else Result := RawNumber;
end;

function TCustomExport.GetDisplayString(var aItems: TStringList; aIndex: Integer): String;
var
  ItemInfo: TExportItemInfo;
begin
  ItemInfo := aItems.Objects[aIndex] as TExportItemInfo;

  // Get Display String
  with ItemInfo do
    if Assigned(ItemInfo) and UseDisplayString then // If we need to use Display String,
      Result := DisplayString       // use it,
    else Result := aItems[aIndex];  // Otherwise use the default string
end;

function TCustomExport.GetDateFormatString: String;
{ Essentially we want to remove all non-(month, day, year, dateseparator) characters from
  the ShortDateFormat String. The only language known to have any is Bulgarian, but that
  can always change in the future }
const
//  cValidDateChars = 'mdy';
  cValidDateChars = 'mdy';
begin
//  Result := StripInvalidChars(LowerCase(ShortDateFormat), cValidDateChars + DateSeparator);
  Result := 'dd/mm/aaaa';//StripInvalidChars(LowerCase(ShortDateFormat), cValidDateChars + DateSeparator);
end;

function TCustomExport.GetTimeFormatString: String;
// Used to get Time Format Strings for Excel & SYLK Exports
begin
  with FOwner do
  begin
    Result := 'hh:mm';                                                // Add Hours and Minutes
    if TimeDisplaySeconds in FOptions then Result := Result + ':ss'; // Add Seconds (if desired)
    if (not (Time24HourFormat in FOptions)) and     // If not 24hour Format Selected,
       Has(TimeAMString) and Has(TimePMString) then // and current Locale has AM/PM Strings then
      Result := Result + '\ AM/PM'; // Add AM/PM Specifier (if desired)
      // Note: Although it says 'AM/PM', Excel will localize the string automatically
  end;
end;

function TCustomExport.GetDateTimeFormatString: String;
// Used to get DateTime Format Strings for Excel & SYLK Exports
begin
  Result := GetDateFormatString + ' ' + GetTimeFormatString;
end;

function TCustomExport.GetCurrencyFormatString: String;
const
  cNumberFormat = '#,##0.%s';
var
  PS, NS, Decimals: String;
  i: Integer;
begin
  // Construct Decimals String Representation (ie. for 2 decimals it should be '00')
  Decimals := '';
  for i := 1 to CurrencyDecimals do
    Decimals := Decimals + '0';

  // Calculate Positive Number String
  case CurrencyFormat of
    0: PS := Format('"%s"' + cNumberFormat,  [CurrencyString, Decimals]); // '$1' Format
    1: PS := Format(cNumberFormat + '"%s"',  [Decimals, CurrencyString]); // '1$' Format
    2: PS := Format('"%s "' + cNumberFormat, [CurrencyString, Decimals]); // '$ 1' Format
    3: PS := Format(cNumberFormat + '" %s"', [Decimals, CurrencyString]); // '1 $' Format
  end;

  // Calculate Negative Number String
  case NegCurrFormat of
    0: NS := Format('\("%s"' + cNumberFormat + '\)', [CurrencyString, Decimals]); // '($1)' Format
    1: NS := Format('"-%s"' + cNumberFormat, [CurrencyString, Decimals]); // '-$1' Format
    2: NS := Format('"%s-"' + cNumberFormat, [CurrencyString, Decimals]); // '$-1' Format
    3: NS := Format('"%s"' + cNumberFormat + '\-', [CurrencyString, Decimals]); // '$1-' Format
    4: NS := Format('\(' + cNumberFormat + '"%s"\)', [Decimals, CurrencyString]); // '(1$)' Format
    5: NS := Format('\-' + cNumberFormat + '"%s"', [Decimals, CurrencyString]); // '-1$' Format
    6: NS := Format(cNumberFormat + '"-%s"', [Decimals, CurrencyString]); // '1-$' Format
    7: NS := Format(cNumberFormat + '"%s-"', [Decimals, CurrencyString]); // '1$-' Format
    8: NS := Format('\-' + cNumberFormat + '" %s"', [Decimals, CurrencyString]); // '-1 $' Format
    9: NS := Format('"-%s "' + cNumberFormat, [CurrencyString, Decimals]); // '-$ 1' Format
   10: NS := Format(cNumberFormat + '" %s-"', [Decimals, CurrencyString]); // '1 $-' Format
   11: NS := Format('"%s "' + cNumberFormat + '\-', [CurrencyString, Decimals]); // '$ 1-' Format
   12: NS := Format('"%s -"' + cNumberFormat, [CurrencyString, Decimals]); // '$ -1' Format
   13: NS := Format(cNumberFormat + '"- %s"', [Decimals, CurrencyString]); // '1- $' Format
   14: NS := Format('\("%s "' + cNumberFormat + '\)', [CurrencyString, Decimals]); // '($ 1)' Format
   15: NS := Format('\(' + cNumberFormat + '" %s"\)', [Decimals, CurrencyString]); // '(1 $)' Format
  end;

  // Construct Currency Format String
  Result := Format('%s_);[Red]%s', [PS, NS]); // Calc. Actual Format String
end;

function TCustomExport.ConvertToDate(S: String): TDateTime;
const
  cNumbers = '0123456789 ';
begin
  Result := StrToDate(StripInvalidChars(S, cNumbers + DateSeparator));
end;

function TCustomExport.ConvertToDateTime(S: String): TDateTime;
{ This is not perfect but it should work. It might not work properly if there is
  an invalid character in the string which is also a character of the TimeAMString
  or TimePMString, but just in the wrong spot in the DateTime String - no matter,
  there are currently no known Languages like this (According to Control Panel/Regional) }
const
  cValidChars = '0123456789 ';
begin
  Result := StrToDateTime(StripInvalidChars(S, cValidChars + DateSeparator +
                                               TimeSeparator + TimeAMString + TimePMString));
end;

function TCustomExport.Decimalize(S: String): String;
var
  i: Integer;
begin
  Result := S;
  if DecimalSeparator <> '.' then
  begin
    i := Pos(DecimalSeparator, Result); // Get Position of Decimal Character (if any)
    if i > 0 then Result[i] := '.';     // Replace Decimal Separator with an actual Decimal
  end;
end;

{$HINTS OFF}
function TCustomExport.IsNumber(S: String): Boolean;
var
  TestFloat: Extended;
  ErrorCode: Integer;
begin
  Val(Decimalize(S), TestFloat, ErrorCode);  // Try to Convert to Floating Point Value
  Result := ErrorCode = 0;                   // See if it's a Valid Number

{ However, in Locales where the Decimal Separater is not '.' the Result would still
  be True at this point, so we must clarify if the string contains a '.' place if
  it is not supposed to and if it does, it must be disqualified as a number since
  FloatToStr will only work with the proper Decimal Separator }

  if Result and (DecimalSeparator <> '.') and (Pos('.', S) > 0) then
    Result := False;
end;

{ There is a remote possibility that this maybe useful sometime in the future...

function TCustomExport.IsInteger(S: String): Boolean;
var
  TestInt: Integer;
  ErrorCode: Integer;
begin
  Val(Decimalize(S), TestInt, ErrorCode); // Try to Convert to Integer Point Value
  Result := ErrorCode = 0;
end;}
{$HINTS ON}

function TCustomExport.GetColInfo(Col: Integer): TExportColumnInfo;
begin
  Result := nil; // assume failure

  with FOwner.FExportColumns do
    if Assigned(Objects[Col]) then
      Result := TExportColumnInfo(Objects[Col]);
end;

function TCustomExport.GetColType(i: Integer; StringValue: String): TColType;
begin
  Result := ctUnknown; // assume failure

  // If we are exporting the Column Headers, make sure the ColType is set to ctString!
  with FOwner do
    if (FNumEntriesExported = 0) and (ShowColHeaders in FOptions) then
      Result := ctString
    else // Otherwise, Calculate the ColType
    begin
      // First See if the Column has an attached ColType
      with FExportColumns do
        if Assigned(Objects[i]) then
          Result := TExportColumnInfo(Objects[i]).ColType;

      // If Result is still ctUknown then try and figure out what it is...
      if Result = ctUnknown then
        if (i = 0) and (NumberRows in FOptions) then // If it's the Row Numbering Column
          Result := ctNumber // Set to Number

        else // otherwise, try to figure it out based on it's string representation
        begin
          if IsNumber(StringValue) then
            Result := ctNumber;

          // In the future, might be good to include IsDate, IsTime and IsDateTime routines
          // here for non-db data to autodetect those data types.
        end;
    end;
end;

function TCustomExport.GetColWidth(i: Integer): Integer;
const
  cMaxDigits = 7;
var
  MaxDigits: Integer;
begin
  with FOwner do
  begin
    // If Numbering Rows...
    if NumberRows in FOptions then
    begin
      if i = 0 then // if Row Number Column
      begin
        // Get Length of Row Number Col Header
        Result := Length(cNumberRowsCaption);

        // Get Number of Items to Export
        MaxDigits := GetNumExportableItems;         // Get Number of Exportable Items,
        if MaxDigits > 0 then                       // if it is not Unknown (-1),
          MaxDigits := Length(IntToStr(MaxDigits)); // then find out the max number of digits we need

        // Make Sure we have enough space for the digits
        if Result < MaxDigits then
          Result := MaxDigits;
        Exit;
      end;

      Dec(i); // Compensate for the Extra Column
    end;

    // Get Column Width
    if (i < FColWidths.Count) and (not IsBlankString(FColWidths[i])) then // If User assigned a col width,
      Result := StrToInt(FColWidths[i])  // then use it, otherwise
    else Result := GetDefaultColWidth(i); // Get Default Column Width
  end;
end;

function TCustomExport.GetTempFile(FileExt: String): String;
const
  cJustInCase  = 'temp.';     // If GetTempHTMLFile fails, it returns this file name

  function RemoveExt(FileName: String): String;
  var
    Ext: String;
    Loc: Integer;
  begin
    Result := '';
    Ext := ExtractFileExt(FileName);
    if Length(Ext) > 0 then
    begin
      Loc := Pos(Ext, FileName);
      Result := Copy(FileName, 0, Loc-1)
    end
    else Result := FileName;
  end;

  function ChangeExt(FileName, NewExtension: String): String;
  begin
    Result := RemoveExt(FileName) + '.' + NewExtension;
  end;

const
  cTempExt = 'TMP';
var
  Buff: array[1..cMaxPathSize] of Char;
begin
  // Create Temp File Name
  if GetTempFileName(PChar(ExpandTempDir('')), PChar(cTempFilePrefix), 0, @Buff) <> 0 then // if got temp file
  begin
    Result := pchar(@Buff);
    DeleteFile(Result); // Erase .TMP file that Windows Creates
    Result := ChangeExt(Result, FileExt); // Change to HTML Extension
  end
  else Result := ExpandTempDir(cJustInCase + FileExt);

  // Add Temp File to TempFiles Stringlist so that it can be deleted when the App Terminates
  with TempFiles do
  begin
    Add(Result);                      // Add Temporary File
    Add(ChangeExt(Result, cTempExt)); // Add Temporary File with .TMP Extension
  end;
end;

function TCustomExport.GetTimeStamp: String;
var
  S: String;
begin
  if Has(cCreatedMsg) then
    S := cCreatedMsg + ' %s'
  else S:= '%s';

  Result := Format(S, [DateTimeToStr(Now)]);
end;

procedure TCustomExport.SetInitialMaxColWidths(var S: TStringList);
var
  i: Integer;
begin
  // Initialize each element of array
  for i := 0 to S.Count - 1 do
    PMaxColWidths^[i] := 0;

  // Set the first max col widths
  UpdateMaxColWidths(S);
end;

procedure TCustomExport.UpdateMaxColWidths(var S: TStringList);
var
  i: Integer;
begin
  // Update Maximum Col Widths for Each column
  for i := 0 to S.Count -1 do
    if Length(S[i]) > PMaxColWidths^[i] then
      PMaxColWidths^[i] := Length(S[i]);
end;

procedure TCustomExport.WriteHeader(var Columns: TStringList);
begin
  if Assigned(FProgForm) then
    FProgForm.Show;

  // Copy Columns so we can reference them at any time in the export
  FColumns.Assign(Columns);

  // Allocate Max Col Widths Array
  PMaxColWidths := AllocMem(FColumns.Count * SizeOf(Integer));

  // Set Maximum Col Widths
  SetInitialMaxColWidths(Columns);
end;

procedure TCustomExport.WriteEntry(var Items: TStringList; Entry: Integer);
begin
  FNumEntriesExported := Entry; // Set Number of Entries Exported
  UpdateMaxColWidths(Items);    // Set Maximum Col Widths
end;

procedure TCustomExport.WriteFooter;
var
  n: Single;
begin
  // Call OnWriteFooter Event with the number of Rows that has been exported
  with FOwner do
    if Assigned(FOnWriteFooter) then
      FOnWriteFooter(FOwner, FCount.RowsExported);

  // Insert the number of RowsExported into the Footer Text
  with FOwner do
  begin
    n := FCount.RowsExported; // Convert to Float
    FFormatFooter.Text := Format(FFooter.Text, [n]);
  end;
end;

procedure TCustomExport.HandleError(Exception: TObject);
begin
  MessageDlg(cGenericExportError, mtError, [mbOk], 0);
end;

function TCustomExport.DoExport: Boolean;
var
  CurrentForm: TForm;
  S: TAutoCleanupStringList;
begin
  Result := False; // default to operation failed
  CurrentForm := nil;
  Application.ProcessMessages;

  // Abort if no data
  if not FOwner.HasData then Exit;

  // Don't Allow Export if User already cancelled
  if FUserCancelled then Exit;

  // Prepare Export Columns
  with FOwner do
  begin
    GetExportColumns(FExportColumns); // Get Export Columns
    if NumberRows in FOptions then // Insert a Number Column
      FExportColumns.Insert(0, cNumberRowsCaption);
  end;

  // Write Header & Entries
  Screen.Cursor := crHourglass; // Set Busy Cursor
  with FOwner do
  try
    // Disable Form
    CurrentForm := TForm(FOwner.Owner);
    if (ShowProgress in FOptions) and Assigned(CurrentForm) then
      CurrentForm.Enabled := False;

    InitExport;
    try
      S := TAutoCleanupStringList.Create;

      WriteHeader(FExportColumns); // Call Descendant's WriteHeader Method

      // Write Each Entry
      InitExportItems; // Prepare for Export
      while MoreExportItems do
      begin
        // Write Export Item
        GetNextExportItem(TStringList(S));               // Get Export Item Data
        WriteEntry(TStringList(S), FCount.RowsExported); // Call Descendant's WriteEntry Method

        // Update Progress Bar
        if Assigned(FProgForm) and (not FProgForm.UpdateProgress(FCount.RowsExported)) then // Exit if User Cancels
        begin
          Result := False;
          Exit;
        end;
      end;

      // Write Footer
      WriteFooter; // Call Descendant's WriteFooter Method

      // Return Successful
      Result := True;

    except
      on Exception do
      begin
        HandleError(ExceptObject); // Call Descendant's Exception Handler

        // Free Progress Form
        if Assigned(FProgForm) then
          FProgForm.Free;
      end;
    end;

  finally
    S.AutoFree;
    CleanUpExport;

    // Re-enable Form
    if (ShowProgress in FOptions) and Assigned(CurrentForm) then
      CurrentForm.Enabled := True;

    Screen.Cursor := crDefault; // Restore Default Cursor
    FExportSuccessful := Result; // Set Export Successful Property
  end;
end;

// *** TCustomShellViewExport Methods ***

constructor TCustomShellViewExport.Create(aOwner: TExportXComponent; aExportFile: String; aExportType: TExportType);
var
  S: String;
  i: Integer;
begin
  inherited;

  // Calculate Export Type String (ie. Microsoft Word, HTML or what?)
  S := FExtension;
  if Length(S) > 0 then
    for i := Integer(Low(cExportTypeExtensions)) to Integer(High(cExportTypeExtensions)) do
      if cExportTypeExtensions[TExportType(i)] = S then
        S := cExportTypes[TExportType(i)];

  // Assign Default Error Message
  ShellViewErrorMsg := Format(cDefShellViewError, [S, S]);

  // Figure out which file to create
  if FExportToFile then                  // If we're exporting to a file,
    FOutputFile := aExportFile            // use user-defined file name
  else FOutputFile := GetTempFile(FExtension); // else, Get Temporary File Name
end;

destructor TCustomShellViewExport.Destroy;
var
  Handle: HWnd;
begin
  // View the File (if ExportToFile Specified)
  if FCreatedSuccessfully and not FExportToFile and FExportSuccessful then
  begin
    Handle := ShellExecute(0, 'open', PChar(FOutputFile), nil, nil, SW_SHOWNORMAL);
    if Handle > 32 then
      FViewerHandle := Handle
    else MessageDlg(ShellViewErrorMsg, mtError, [mbOk], 0); // Show Error Message (assumes the user has no viewer installed and that's the reason for the error)
  end;

  inherited;
end;

// *** TCustomTextExport Methods ***

constructor TCustomTextExport.Create(aOwner: TExportXComponent; aExportFile: String; aExportType: TExportType);
begin
  inherited;

  // Initialize
  FCreatedSuccessfully := False;
  FAlreadyClosed := False;

  // Create the Text File
  AssignFile(F, FOutputFile);
  Rewrite(F);

  FCreatedSuccessfully := True;
end;

destructor TCustomTextExport.Destroy;
begin
  // Close the Text File
  if FCreatedSuccessfully then
    CloseTextFile;

  inherited;
end;

procedure TCustomTextExport.CloseTextFile;
begin
  if not FAlreadyClosed then
  begin
    CloseFile(F);
    FAlreadyClosed := True;
  end;
end;

procedure TCustomTextExport.HandleError(Exception: TObject);
begin
  MessageDlg(cTextFileError, mtError, [mbOk], 0);
end;

// *** TClipboardExport Methods ***

procedure TClipboardExport.WriteHeader(var Columns: TStringList);
begin
  FClipboardText := GetPaddedHeader(Columns);
end;

procedure TClipboardExport.WriteEntry(var Items: TStringList; Entry: Integer);
begin
  FClipboardText := FClipboardText + GetPaddedEntry(Items, Entry);
end;

procedure TClipboardExport.WriteFooter;
var
  S: String;
begin
  inherited;

  S := GetPaddedFooter;
  if Length(S) > 0 then
    FClipboardText := FClipboardText + S;
end;

destructor TClipboardExport.Destroy;
begin
  if FCreatedSuccessfully then
  begin
    // Copy to Clipboard
    Clipboard.AsText := FClipboardText;

    // Inform User that data was exported successfully.
    if ShowClipboardMsg in FOwner.FOptions then
    begin
      if Assigned(FProgForm) then FProgForm.Hide; // Hide Progress Bar Form
      MessageDlg(cClipSuccessfulMsg, mtInformation, [mbOk], 0); // Show User Message
    end;
  end;

  inherited;
end;

// *** TDetailedClipboardExport Methods ***

procedure TDetailedClipboardExport.WriteHeader(var Columns: TStringList);
begin
  inherited;

  CalculateMaxColumnLength;  

  S := ''; // Initialize

  with FOwner do
  begin
    if Has(FTitle) then
      S := S + FTitle + CRLF + GetUnderline(Length(FTitle)) + CRLFx2;

    if TimeStamp in FOptions then
      S := S + GetTimeStamp + CRLFx2;

    if Has(FHeader.Text) then
      S := S + FHeader.Text + CRLFx2;
  end;
end;

procedure TDetailedClipboardExport.WriteEntry(var Items: TStringList; Entry: Integer);
var
  i: Integer;
  SRow: String;
begin
  inherited;

  with FOwner do
    for i := 0 to Items.Count - 1 do // Write Each Field
    begin
      // Get Column Header
      if ShowColHeaders in FOptions then
      begin
        if Has(FColumns[i]) then
          SRow := Pad(FColumns[i], FMaxColumnLength, False)  + ': '
        else SRow := '';
      end
      else SRow := ''; // If we're only exporting the data.

      S := S + SRow + Strip(Items[i]) + CRLF;
    end;

  // Write a Spacing Line
  S := S + CRLF;
end;

procedure TDetailedClipboardExport.WriteFooter;
begin
  inherited;

  with FOwner do
    if Has(FFormatFooter.Text) then
    begin
      // If Single Column Export, Write a blank line first
      if GetNumDataColumns(FColumns.Count, FOptions) = 1 then
        S := S + CRLF;

      // Write Footer
      S := S + FFormatFooter.Text;
    end;
end;

destructor TDetailedClipboardExport.Destroy;
begin
  if FCreatedSuccessfully then
  begin
    // Copy to Clipboard
    Clipboard.AsText := S;

    // Inform User that data was exported successfully.
    if ShowClipboardMsg in FOwner.FOptions then
    begin
      if Assigned(FProgForm) then FProgForm.Hide; // Hide Progress Bar Form
      MessageDlg(cClipSuccessfulMsg, mtInformation, [mbOk], 0); // Show User Message
    end;
  end;

  inherited;
end;

// *** TBIFFExport Methods ***

constructor TBIFFExport.Create(aOwner: TExportXComponent; aExportFile: String; aExportType: TExportType);
  function GetUserLocale: Word;
  const
    cInternationalKey = 'Control Panel\International';
    cLocaleValue      = 'Locale';
  begin
     Result := $0000; // Initialize

    with TMyReg.Create do
    try
      if OK_ReadOnly(cInternationalKey) then
      begin
        Result := StrToInt('$' + ReadString(cLocaleValue)); // Read the Hex Locale Value
        CloseKey;
      end;

    finally
      Free;
    end;
  end;

begin
  inherited;

  // Set Biff Version (Can be 2, 3, 4 or 5)
  case GetUserLocale of
    $0C09, $1009, $2409, $1809, $2009, $1409, $1C09, $0809, $0409, // English Locales
    $0404,                    // Chinese (Taiwan)
    $140C,                    // French (Luxembourg)
    $1007: BIFF_Version := 2; // German (Luxembourg)
  else
    BIFF_Version := 5; // otherwise, if other Locale
  end;

  // Create Output File
  AssignFile(F, FOutputFile);
  Rewrite(F, 1); // Create File with Record Size of 1
end;

procedure TBIFFExport.WriteRecordHeader;
const
 LEN_RECORDHEADER = SizeOf(Word) * 2;
var
 awBuf : array[0..1] of word;
begin
  awBuf[0] := RecType;
  awBuf[1] := RecLength;
  BlockWrite(F, awBuf, LEN_RECORDHEADER);
end;

procedure TBIFFExport.WriteByteRecord(aRecType: Word; aByte: Byte);
begin
  // Write Record Header
  RecType   := aRecType;
  RecLength := SizeOf(Byte);
  WriteRecordHeader;

  // Write Byte Value
  BlockWrite(F, aByte, RecLength);
end;

{procedure TBIFFExport.WriteWordRecord(aRecType, aWord: Word);
begin
  // Write Record Header
  RecType   := aRecType;
  RecLength := SizeOf(Word);
  WriteRecordHeader;

  // Write Word Value
  BlockWrite(F, aWord, RecLength);
end;}

procedure TBIFFExport.WriteFlagRecord(aRecType: Word; Flag: Boolean);
const
  cFlagged   = $01;
  cNotFlagged = $00;
var
  aByte: Byte;
begin
  if Flag then
    aByte := cFlagged
  else aByte := cNotFlagged;

  WriteByteRecord(aRecType, aByte);
end;

procedure TBIFFExport.WriteStringRecord(aRecType: Word; S: String);
var
  Len: Byte;
begin
  // Write Record Header
  RecType   := aRecType;
  RecLength := Length(S) + 1;
  WriteRecordHeader;

  // Write Format String
  S := trim(Copy(S, 1, 255));      // Truncate to 255 chars before writing
  Len := Length(S);                // Get Length of String
  BlockWrite(F, Len, SizeOf(Len)); // Write Length of String
  BlockWrite(F, PansiChar(Ansistring(S))^, Len);   // Write String
end;

procedure TBIFFExport.WriteBOF;
const
  DOCTYPE_XLS = $0010;

  BIT_BIFF3 = $0200;
  BIT_BIFF4 = $0400;
  BIT_BIFF5 = $0800;

  BOF_BIFF3 = BIFF_BOF or BIT_BIFF3;
  BOF_BIFF4 = BIFF_BOF or BIT_BIFF4;
  BOF_BIFF5 = BIFF_BOF or BIT_BIFF5;

var
  awBuf : array[0..2] of word;
begin
  awBuf[0] := 0;
  awBuf[1] := DOCTYPE_XLS;
  awBuf[2] := 0;

  case BIFF_Version of
    2: begin RecType := BIFF_BOF;  RecLength := 4; end;
    3: begin RecType := BOF_BIFF3; RecLength := 6; end;
    4: begin RecType := BOF_BIFF4; RecLength := 6; end;
    5: begin RecType := BOF_BIFF5; RecLength := 6; end;
  end;

  WriteRecordHeader;
  BlockWrite(F, awBuf, RecLength);
end;

procedure TBIFFExport.WriteEOF;
begin
  RecType   := BIFF_EOF;
  RecLength := 0;
  WriteRecordHeader;
end;

procedure TBIFFExport.WriteDimensions(MaxRows, MaxCols: Integer);
const
  BIT_BIFF3  = $0200;
  DIMENSIONS_BIFF3 = BIFF_DIMENSIONS or BIT_BIFF3;

var
  awBuf : array[0..4] of Word;
begin
  awBuf[0] := 0;
  awBuf[1] := MaxRows;
  awBuf[2] := 0;
  awBuf[3] := MaxCols;
  awBuf[4] := 0;

  if BIFF_Version = 2 then
  begin
    RecType := BIFF_DIMENSIONS;
    RecLength := 8;
  end
  else
  begin
    RecType := DIMENSIONS_BIFF3;
    RecLength := 10;
  end;

  DimensionsOffset := FilePos(F); // Save Offset of Dimensions Record
  WriteRecordHeader;
  BlockWrite(F, awBuf, RecLength);
end;

procedure TBIFFExport.WriteXF(aParams: array of Byte); // Write Extended Format Record
const
  cNumParams        = 3;
  cFormatIndex      = 0;
  cFontIndex        = 1;
  cBorderAlignIndex = 2;

  cLEN_DATA = SizeOf(Byte) * 4;

var
 awBuf : array[0..3] of Byte;
begin
  // Make Sure Proper Number of Parameters have been passed
  if SizeOf(aParams) <> 3 then
    raise Exception.Create('TBIFFExport.WriteXF: Invalid Number of Parameters');

  // Write Record Header
  RecType   := BIFF_XF;
  RecLength := 4;
  WriteRecordHeader;

  // Write Data
  awBuf[0] := BIFF_BaseFontIndex + aParams[cFontIndex] ; // Write Font Index #
  awBuf[1] := 0; // ?
  awBuf[2] := BIFF_BaseFormatIndex + aParams[cFormatIndex]; // Write Cell Format String Index
  awBuf[3] := aParams[cBorderAlignIndex];                   // Write Border and Alignment Info
  BlockWrite(F, awBuf, cLEN_DATA);
end;

procedure TBIFFExport.WriteFont(aFontName: String; aHeight: Integer; aFontStyles: TFontStyles);
const
  cBoldBit      = $01;
  cItalicBit    = $02;
  cUnderlineBit = $04;
  cStrikeOutBit = $08;

type
  TFontRec = record
    FontHeight: Word;
    FontAttributes,
    FontReserved,     // Reserved Byte - just set to $00
    FontNameLength: Byte;
  end;

const
  LEN_TFONTREC = SizeOf(TFontRec) - 1; // Align to Even Value

var
 FontRec: TFontRec;
begin
  // Populate the Font Record
{  with FontRec do
  begin
    FontHeight   := aHeight * 20; // Font Height is in 1/20ths of a point
    FontReserved := $00;        // Reserved Byte

    // Set Font Styles
    FontAttributes := $00; // Initialize
    if fsBold      in aFontStyles then FontAttributes := FontAttributes or cBoldBit;
    if fsItalic    in aFontStyles then FontAttributes := FontAttributes or cItalicBit;
    if fsUnderline in aFontStyles then FontAttributes := FontAttributes or cUnderlineBit;
    if fsStrikeOut in aFontStyles then FontAttributes := FontAttributes or cStrikeOutBit;

    // Set Length of Font Name
    FontNameLength := Length(aFontName);
  end;

  // Write Record Header
  RecType   := BIFF_FONT;
  RecLength := LEN_TFONTREC + FontRec.FontNameLength;
  WriteRecordHeader;

  // Write Data
  BlockWrite(F, FontRec, LEN_TFONTREC);                    // Write Font Information
  BlockWrite(F, PChar(aFontName)^, FontRec.FontNameLength); // Write Font Name String
}
end;

{
// Unsuccessful attempt at a BIFF5 Font Record
procedure TBIFFExport.WriteFont(aFontName: String; aHeight: Integer; aFontStyles: TFontStyles);
const
  cBoldBit      = $01;
  cItalicBit    = $02;
  cUnderlineBit = $04;
  cStrikeOutBit = $08;

type
  TFontRec = record       // [BIFF5 0231h Reccord Names]
    FontHeight,           // dyHeight
    FontAttributes,       // grbit
    ColorPalette,         // icv
    BoldStyle,            // bis
    SuperSubScript: Word; // sss
    Underline,            // uls
    FontFamily,           // bFamily
    CharSet,              // bCharSet
    FontReserved,         // (Reserved)
    FontNameLength: Byte; // cch
  end;

const
  LEN_TFONTREC = SizeOf(TFontRec);

var
 FontRec: TFontRec;
begin
  // Populate the Font Record
  with FontRec do
  begin
    // Initialize
    FontReserved   := $00; // Reserved Byte
    SuperSubScript := $00;
    FontFamily     := FF_DONTCARE;
    CharSet        := DEFAULT_CHARSET;

    // Font Height is in 1/20ths of a point
    FontHeight := aHeight * 20;

    // Set Font Styles
    FontAttributes := $FF00; // Initialize
//    if fsBold      in aFontStyles then FontAttributes := FontAttributes or cBoldBit;
//    if fsItalic    in aFontStyles then FontAttributes := FontAttributes or cItalicBit;
//    if fsUnderline in aFontStyles then FontAttributes := FontAttributes or cUnderlineBit;
//    if fsStrikeOut in aFontStyles then FontAttributes := FontAttributes or cStrikeOutBit;

    ColorPalette := $0000; // Default Color (Black)
    BoldStyle := $0000;
    Underline := $02;

    // Set Length of Font Name
    FontNameLength := Length(aFontName);
  end;

  // Write Record Header
  RecType   := BIFF5_FONT;
  RecLength := LEN_TFONTREC + FontRec.FontNameLength;
  WriteRecordHeader;

  // Write Data
  BlockWrite(F, FontRec, LEN_TFONTREC);                    // Write Font Information
  BlockWrite(F, PChar(aFontName)^, FontRec.FontNameLength); // Write Font Name String
end;}

procedure TBIFFExport.WriteFormat(aFormatString: String);
begin
  WriteStringRecord(BIFF_FORMAT, aFormatString);
end;

procedure TBIFFExport.WritePrintHeader(S: String);
begin
  WriteStringRecord(BIFF_HEADER, S);
end;

procedure TBIFFExport.WritePrintFooter(S: String);
begin
  WriteStringRecord(BIFF_FOOTER, S);
end;

procedure TBIFFExport.WriteColWidth(aColIndex: Byte; NumCharacters: Word);
type
  TColWidthRec = record
    FirstCol,
    LastCol: Byte;
    Width: Word;
  end;

var
  ColWidthRec: TColWidthRec;

begin
  // Write the Record Header
  RecType   := BIFF_COLWIDTH;
  RecLength := 4;
  WriteRecordHeader;

  // Prepare the Data
  with ColWidthRec do
  begin
    FirstCol := aColIndex;
    LastCol  := aColIndex;
    Width    := NumCharacters * 256; // Because ColWidths are in units of 1/256 of a character
  end;

  // Write the Data
  BlockWrite(F, ColWidthRec, SizeOf(TColWidthRec));
end;

procedure TBIFFExport.WriteData(ColType: TColType; ARow, ACol: Integer; AData: Pointer);
const
  cPadding: Byte = $00; // This is unused. In older versions it was the byte that
                       // specified the borders and alignments

var
  awBuf : array[0..2] of word;
  AWordLength: Word;
  ABoolByte: Byte;
  BIFF_Fmt: Byte;
begin
  // Set Format
  with FOwner do
    if ARow = FTitleRow      then BIFF_Fmt := BIFF_TitleFormat     else // Title Format
    if ARow = FTimeStampRow  then BIFF_Fmt := BIFF_TimeStampFormat else // TimeStamp Format
    if ARow = FHeaderRow     then BIFF_Fmt := BIFF_HeaderFormat    else // Header Format
    if ARow = FFooterRow    then BIFF_Fmt := BIFF_FooterFormat     else // Footer Format
    if (ARow = FColHeadersRow) and (ShowColHeaders in FOptions) then BIFF_Fmt := BIFF_ColHeaderFormat   // Col Headers Format
    else BIFF_Fmt := BIFF_GeneralFormat; // else use the General Format

  // Write Pre-Rec Info Structure Data
  case ColType of
    ctNumber: WriteNumber;

    ctCurrency:
    begin
      BIFF_Fmt := BIFF_CurrencyFormat;
      WriteNumber;
    end;

    ctDate, ctDateTime_ShowDateOnly:
      begin
        BIFF_Fmt := BIFF_DateFormat;
        WriteNumber;
      end;

    ctTime, ctDateTime_ShowTimeOnly:
      begin
        BIFF_Fmt := BIFF_TimeFormat;
        WriteNumber;
      end;

    ctDateTime:
      begin
//        if length(pchar(AData)) > 10 then
//        if Pansichar(ansiString(AData))>10 then
           BIFF_Fmt := BIFF_DateTimeFormat;
//        else
//           BIFF_Fmt := BIFF_DateFormat;
        WriteNumber;
      end;

    ctBoolean : WriteBoolean;

  else        //aqui
    WriteLabel(AWordLength, PansiChar(ansiString(AData))); // else, Assume it's a String
  end;

  // Write Record Info Structure
  awBuf[0] := ARow;
  awBuf[1] := ACol;
  awBuf[2] := BIFF_BaseFormatIndex + BIFF_Fmt + 1;
  BlockWrite(F, awBuf, SizeOf(awBuf)); // Write Row, Column, ixfe Information
  BlockWrite(F, cPadding, SizeOf(cPadding));

  // Write Data
  case ColType of
    ctNumber,    // Don't write anything for these ColTypes
    ctCurrency,
    ctDate,
    ctTime,
    ctDateTime,
    ctDateTime_ShowDateOnly,
    ctDateTime_ShowTimeOnly:;

    ctBoolean:
      begin
        if Byte(AData^) <> 0 then ABoolByte := 1 else ABoolByte := 0;
        BlockWrite(F, ABoolByte, SizeOf(ABoolByte));
        ABoolByte := 0;
        BlockWrite(F, ABoolByte, SizeOf(ABoolByte));
      end;

  else // otherwise, assume it's a string
    begin
      ABoolByte := AWordLength;
      BlockWrite(F, ABoolByte, SizeOf(ABoolByte))
    end;
  end;

  if RecLength <> 0 then BlockWrite(F, AData^, RecLength);
end;

{procedure TBIFFExport.WriteBlank;
begin
  RecType   := BIFF_BLANK;
  RecLength := 7;
  WriteRecordHeader;
  RecLength := 0;
end;}

procedure TBIFFExport.WriteNumber;
begin
  RecType   := BIFF_NUMBER;
  RecLength := 15;
  WriteRecordHeader;
  RecLength := 8;
end;

procedure TBIFFExport.WriteLabel(var w: Word; AData: Pointer);
begin
  w         := StrLen(PAnsichar(AData));
  RecType   := BIFF_LABEL;
  RecLength := 8+w;

  WriteRecordHeader;
  RecLength := w;
end;

procedure TBIFFExport.WriteBoolean;
begin
  // Write Record Header
  RecType   := BIFF_BOOLEAN;
  RecLength := 9;
  WriteRecordHeader;
  RecLength := 0;
end;

procedure TBIFFExport.WriteCell(Value: String; Row, Col: Integer;
                                ColType: TColType; ItemInfo: TExportItemInfo);
const
  cMax_XL_Rows = 65536;
  cMax_XL_Cols = 256;

  cMaxLength = 255;

var
  ANumber  : Double;
  ABoolean : Boolean;
begin

  // Abort if out of Range
  if (Row > cMax_XL_Rows) or (Col > cMax_XL_Cols) then exit;

  // Figure out the Data Type
  if ColType = ctUnknown then   // If we are not given a Column Type,
    ColType := GetColType(Col, Value); // try to figure it out

  // Write Data
  case ColType of
    ctNumber:
      if not IsBlankString(Value) then
      begin
        ANumber := StrToFloat(GetRoundedNumber(Value, ItemInfo));
        WriteData(ColType,  Row, Col, @ANumber);
      end;

    ctCurrency:
      if not IsBlankString(Value) then
      begin
        ANumber := StrToFloat(Value);
        WriteData(ColType, Row, Col, @ANumber);
      end;

    ctDate:
      if not IsBlankString(Value) then
      begin
        ANumber := ConvertToDate(Value);
        WriteData(ColType,  Row, Col, @ANumber);
      end;

    ctTime:
      if not IsBlankString(Value) then
      begin
        ANumber := StrToTime(Value);
        WriteData(ColType,  Row, Col, @ANumber);
      end;

    ctDateTime:
      if not IsBlankString(Value) then
      begin
        ANumber := ConvertToDateTime(Value);
        WriteData(ColType,  Row, Col, @ANumber);
      end;

    ctDateTime_ShowDateOnly:
      if not IsBlankString(Value) then
      begin
        ANumber := Int(ConvertToDateTime(Value)); // Get Date Value of DateTime Value
        WriteData(ColType,  Row, Col, @ANumber);
      end;

    ctDateTime_ShowTimeOnly:
      if not IsBlankString(Value) then
      begin
        ANumber := Frac(ConvertToDateTime(Value)); // Get Time Value of DateTime Value
        WriteData(ColType,  Row, Col, @ANumber);
      end;

    ctBoolean:
      with ItemInfo do
      begin
       case BooleanValue of
         bvTrue : ABoolean := True;
         bvFalse: ABoolean := False;
       end;

       // Write Boolean Value only if it has a value assigned to it
       if BooleanValue <> bvNull then
         WriteData(ctBoolean, Row, Col, @ABoolean);
      end;

  else // Otherwise, must be a String Value
    Value := FOwner.Strip(Value); // Remove Control Characters
    if Length(value) > cMaxLength then
       Value := Copy(value, 1, cMaxLength); // Truncate

    WriteData(ColType, Row, Col, PAnsiChar(AnsiString(value)));//PAnsiChar(AnsiString(Value)));
  end;
end;

procedure TBIFFExport.WriteHeader(var Columns: TStringList);
var
  i: Integer;
  HasTitle, HasTimeStamp, HasHeader: Boolean;
begin
  inherited;

  // Write BIFF Header
  WriteBOF;
  WriteDimensions(0, 0); // We don't really know the dimensions yet, we'll write them again later

  // Write Print Options
  with FOwner.FXL_Options do
  begin
    WritePrintHeader(FPageHeader);                    // Write Print Header
    WritePrintFooter(FPageFooter);                    // Write Print Footer
    WriteFlagRecord(BIFF_ROWHEADERS, FPrintHeadings); // Write Print Row Headers Flag
    WriteFlagRecord(BIFF_GRIDLINES, FPrintGridlines); // Write Print Gridlines Flag
  end;

  // Write Font Table
  WriteFont('Arial', 10, []);                    // Write Spreadsheet Default Font,
  WriteFont('Arial', 10, []);                    // Write it 4 Times because...
  WriteFont('Arial', 10, []);                    // ...we are not going to use these ...
  WriteFont('Arial', 10, []);                    // ...and they are for compatibility with Older BIFF versions

  WriteFont('Arial', 10, []);                    // Write our Default Font for Data
  WriteFont('Arial', 10, [fsItalic]);            // Write Italic Font
  WriteFont('Arial', 16, [fsBold, fsUnderline]); // Write Title Font
  WriteFont('Arial', 10, [fsBold]);              // Write ColHeadersRow Font

  // Write Formats
  WriteFormat(BIFF_GeneralFormatString);  // Write General Format
  WriteFormat(GetDateFormatString);       // Write Date Format
  WriteFormat(GetTimeFormatString);       // Write Time Format
  WriteFormat(GetDateTimeFormatString);   // Write Date Time Format
  WriteFormat(GetCurrencyFormatString);   // Write Currency Format

  // Write Extended Format Table
  WriteXF([$00, $00, $00]);               // Not sure why this is necessary. Just copying what Excel does.

  WriteXF([BIFF_GeneralFormat, BIFF_DataFont, $00]);   // General Format Data XF Record
  WriteXF([BIFF_DateFormat, BIFF_DataFont, $00]);      // Date Format Data XF Record
  WriteXF([BIFF_TimeFormat, BIFF_DataFont, $00]);      // Time Format Data XF Record
  WriteXF([BIFF_DateTimeFormat, BIFF_DataFont, $00]);  // DateTime Format Data XF Record
  WriteXF([BIFF_CurrencyFormat, BIFF_DataFont, $00]);  // Currenct Format XF Record

  WriteXF([BIFF_GeneralFormat, BIFF_DetailedTitle, $00]);  // Title Format XF Record
  WriteXF([BIFF_GeneralFormat, BIFF_ItalicFont, $00]); // Italic Format Data XF Record
  WriteXF([BIFF_GeneralFormat, BIFF_ColHeadersFont, BIFF_TopBorder or BIFF_BottomBorder]); // ColHeadersRow Data XF Record

  // Calculate ColumnHeadersRow
  with FOwner do
  begin
    // Initialize
    FTitleRow     := BIFF_RowNotPresent;
    FTimeStampRow := BIFF_RowNotPresent;
    FHeaderRow    := BIFF_RowNotPresent;
    FFooterRow    := BIFF_RowNotPresent;

    HasTitle     := Has(FTitle);
    HasTimeStamp := TimeStamp in FOptions;
    HasHeader    := Has(FHeader.Text);

    FColHeadersRow := 0;
    if HasTitle then
    begin
      FTitleRow := 0;
      Inc(FColHeadersRow);
    end;
    if HasTimeStamp then
    begin
      FTimeStampRow := FColHeadersRow;
      Inc(FColHeadersRow);
    end;
    if HasHeader then
    begin
      if HasTitle or HasTimeStamp then Inc(FColHeadersRow);
      FHeaderRow := FColHeadersRow;
    end;
    if HasTitle or HasTimeStamp or HasHeader then Inc(FColHeadersRow);

    // Insert Rows
    if FColHeadersRow > 1 then
    begin
      // Write Title
      if HasTitle then
        WriteCell(FTitle, FTitleRow, 0, ctString, nil);

      // Write Time Stamp
      if HasTimeStamp then
        WriteCell(GetTimeStamp, FTimeStampRow, 0, ctString, nil);

      // Write User Header Text
      if HasHeader then
        for i := 0 to FHeader.Count - 1 do
        begin
          WriteCell(FHeader[i], FHeaderRow + i, 0, ctString, nil);
          Inc(FColHeadersRow);
        end;
    end;
  end;

  // Write Column Headers
  if ShowColHeaders in FOwner.FOptions then
    for i := 0 to Columns.Count - 1 do
      WriteCell(Columns[i], FColHeadersRow, i, ctString, nil)
  else Dec(FColHeadersRow); // If no column headers, decrease ColHeadersRow offset
end;

procedure TBIFFExport.WriteEntry(var Items: TStringList; Entry: Integer);
var
  i: Integer;
begin
  inherited;

  for i := 0 to Items.Count - 1 do
    WriteCell(Items[i], FColHeadersRow + Entry, i, ctUnknown, TExportItemInfo(Items.Objects[i]));
end;

procedure TBIFFExport.WriteFooter;
const
  cSpacing = 2; // Spacing of n Characters to avoid squishing
var
  i, m: Integer;
begin
  inherited;

  // Write Footer Text
  with FOwner do
    if Has(FFormatFooter.Text) then
    begin
      FFooterRow := FCount.RowsExported + FColHeadersRow + 2; // Calculate First row of Footer
      for i := 0 to FFormatFooter.Count - 1 do
        WriteCell(FFormatFooter[i], FFooterRow + i, 0, ctString, nil);
    end;

  // Write Column Widths
  for i := 0 to FColumns.Count - 1 do
  begin
    m := PMaxColWidths^[i] + cSpacing; // Get Max Col Width for Column
    if m > 255 then m := 255; // Because of Excel's 255 char max limit for strings, we need to do this
    WriteColWidth(i, m);      // Write ColWidth
  end;

  // Close File
  WriteEOF;                  // Write EOF Record
  Seek(F, DimensionsOffset); // Goto Dimensions Record
  WriteDimensions(FOwner.FCount.RowsExported + 2, FColumns.Count + 1); // Write the proper dimensions
  Close(F);
end;

// *** TExcelExport Methods ***
{destructor TExcelExport.Destroy;
begin
  inherited;
end;}

// *** TCSVExport Methods ***

function TCSVExport.MakeCSV(var Items: TStringList): String;
  function GetListSeparator: ShortString;
  const
    cRegKey   = 'Control Panel\International';
    cRegValue = 'sList';
  begin
    with TMyReg.Create do
    try
      RootKey := HKEY_CURRENT_USER;

      Result := '';
      if OK_ReadOnly(cRegKey) then
      begin
        // Get List Separator from Registry
        if ValueExists(cRegValue) then
          Result := ReadString(cRegValue);

        CloseKey;
      end;

      // If we didn't get it sucessfully, let's construct it ourselves
      if Length(Result) = 0 then
      begin
        if DecimalSeparator = ',' then
          Result := ';'
        else Result := ',';
      end;
{      Result := ',';
      if VersaoWindows in [WindowsServer2003,WindowsVista, WindowsSeven] then
         Result := ';'}
    finally
      Free;
    end;
  end;

const
  cQuote = '"';
var
  S, ListSeparator: String;
  i: Integer;
begin
{ Make CSV Line
  - For Numbers, convert all foreign ',' style, etc.. separators to an actual decimal separator
  - Use the ListSeparator value and not the Comma (all of the time) in Germany should be a ';'
  - Don't quote Numbers
}

  ListSeparator := GetListSeparator;
  Result := '';
  for i := 0 to Items.Count - 1 do
  begin
    // Calculate Current Item Value
    if GetColType(i, Items[i]) in cNumericColTypes then
      S := GetRoundedNumber(Items[i], TExportItemInfo(Items.Objects[i]))
    else S := cQuote + ReinforceSymbol(FOwner.Strip(Items[i]), cQuote, cQuote) + cQuote;

    // Add to CSV Line
    Result := Result + S;
    if i <> Items.Count - 1 then          // Add ListSeparator (ie. ',' or ';') Value ...
      Result := Result + ListSeparator;  // ... unless it's the last item
  end;
end;

procedure TCSVExport.WriteHeader(var Columns: TStringList);
begin
  inherited;
  WriteLn(F, FOwner.Title);       //montar cabeçalho no CSV
  WriteLn(F, FOwner.Header.text); //montar cabeçalho no CSV
  // Write the Columns Row (as long as user wants to)
  if ShowColHeaders in FOwner.FOptions then
    WriteLn(F, MakeCSV(Columns));
end;

procedure TCSVExport.WriteEntry(var Items: TStringList; Entry: Integer);
begin
  inherited;

  WriteLn(F, MakeCSV(Items)); // Write the Entry
end;

// *** TTabExport Methods ***

procedure TTabExport.WriteHeader(var Columns: TStringList);
const
  cTab = Chr($09);
var
  S: String;
  i: Integer;
begin
  inherited;

  // Write the Columns Row (as long as user wants to)
  if ShowColHeaders in FOwner.FOptions then
  begin
    S := '';
    for i := 0 to Columns.Count - 1 do
      S := S + Columns[i] + cTab;

    WriteLn(F, S);
  end;
end;

procedure TTabExport.WriteEntry(var Items: TStringList; Entry: Integer);
const
  cTab = Chr($09);
var
  S: String;
  i: Integer;
begin
  inherited;

  S := '';
  with FOwner do
    for i := 0 to Items.Count - 1 do
    begin
      if GetColType(i, Items[i]) in cNumericColTypes then
        S := S + GetRoundedNumber(Items[i], TExportItemInfo(Items.Objects[i]))
      else S := S + Strip(Items[i]);

      // Add Tab (if not last item)
      if i < Items.Count - 1 then
        S := S + cTab;
    end;

  WriteLn(F, S);
end;

// *** TTextExport Methods ***

procedure TTextExport.WriteHeader(var Columns: TStringList);
begin
  Write(F, GetPaddedHeader(Columns));
end;

procedure TTextExport.WriteEntry(var Items: TStringList; Entry: Integer);
begin
  Write(F, GetPaddedEntry(Items, Entry));
end;

procedure TTextExport.WriteFooter;
var
  S: String;
begin
  inherited;

  S := GetPaddedFooter;
  if Length(S) > 0 then
    Write(F, S);
end;


// *** TDetailedTextExport Methods ***

procedure TDetailedTextExport.WriteHeader(var Columns: TStringList);
begin
  inherited;

  CalculateMaxColumnLength;

  with FOwner do
  begin
    if Has(FTitle) then
    begin
      WriteLn(F, FTitle);
      WriteLn(F, GetUnderline(Length(FTitle)));
      WriteLn(F, '');
    end;

    if TimeStamp in FOptions then
    begin
      WriteLn(F, GetTimeStamp);
      WriteLn(F, '');
    end;

    if Has(FHeader.Text) then
    begin
      WriteLn(F, FHeader.Text);
      WriteLn(F, '');
    end;
  end;
end;

procedure TDetailedTextExport.WriteEntry(var Items: TStringList; Entry: Integer);
var
  i: Integer;
  S: String;
begin
  inherited;

  with FOwner do
    for i := 0 to Items.Count - 1 do // Write Each Field
    begin
      if ShowColHeaders in FOptions then
      begin
        if Has(FColumns[i]) then
          S := Pad(FColumns[i], FMaxColumnLength, False)  + ': '
        else S := '';
      end
      else S := ''; // If we're only exporting the data

      WriteLn(F, S + Strip(Items[i]));
    end;

    // Write a Spacing Line
    WriteLn(F, '');
end;

procedure TDetailedTextExport.WriteFooter;
begin
  inherited;

  with FOwner do
    if Has(FFormatFooter.Text) then
      WriteLn(F, FFormatFooter.Text);
end;

// *** TDIFExport Methods ***

procedure TDIFExport.WriteTuple(var Items: TStringList);
const
  cTupleTopic   = '-1,0' + CRLF + 'BOT';
  cNumericTopic = '0,%s' + CRLF + 'V';
  cStringTopic  = '1,0'  + CRLF + '"%s"';
var
  S: String;
  i: Integer;
begin
  inherited;

  WriteLn(TempF, cTupleTopic); // Mark Beginning of Tuple (Row)

  for i := 0 to Items.Count - 1 do // Write Each Field
  begin
    S := FOwner.Strip(Items[i]);
    if GetColType(i, S) in cNumericColTypes then
      WriteLn(TempF, Format(cNumericTopic,
                            [GetRoundedNumber(S, TExportItemInfo(Items.Objects[i]))])) // write as numeric data
    else WriteLn(TempF, Format(cStringTopic, [ReinforceSymbol(S, '"', '"')])); // Write String Data
  end;
end;

procedure TDIFExport.WriteHeader(var Columns: TStringList);
const
  cUniqueExt    = 'z93'; // Not necessarily unique but we need it to be different than the DIF
                         // extension since we sometimes try to get 2 temporary file names almost
                         // at the same time (one for F and another for TempF) and windows will
                         // actually sometimes return the same temporary file name if we use the same
                         // extension both times!

  cTableTopic   = 'TABLE'   + CRLF + '0,1'  + CRLF + '"%s"';
  cVectorsTopic = 'VECTORS' + CRLF + '0,%d' + CRLF + '""';
  cTuplesTopic  = 'TUPLES'  + CRLF + '""'; // We don't put the count here, we add it in WriteFooter
  cDataTopic    = 'DATA'    + CRLF + '0,0'  + CRLF + '""';
begin
  inherited;

  // Create the Temporary File
  FTempDIF := GetTempFile(cUniqueExt); // Get Temp File Name w/different extension than DIF
  AssignFile(TempF, FTempDIF);
  Rewrite(TempF);

  // Write DIF Header Information
  with FOwner do
  begin
    WriteLn(TempF, Format(cTableTopic, [FTitle]));              // Write Table Topic
    WriteLn(TempF, Format(cVectorsTopic, [Columns.Count]));     // Write Vectors Topic
    WriteLn(TempF, cTuplesTopic);                               // Write Tuples Topic
    WriteLn(TempF, cDataTopic);                                 // Write Data Topic

    if ShowColHeaders in FOptions then // If we export column headers,
      WriteTuple(Columns);             // Write the Column Headers Row
  end;
end;

procedure TDIFExport.WriteEntry(var Items: TStringList; Entry: Integer);
begin
  inherited;

  WriteTuple(Items); // Write the Tuple (Row)
end;

procedure TDIFExport.WriteFooter;
const
  cEndDataTopic    = '-1,0' + CRLF + 'EOD';
  cTuplesCountLine = 8;
var
  S: String;
  Line, nTuple: Integer;
begin
  inherited;

  WriteLn(TempF, cEndDataTopic); // Write End of Data Topic

  CloseFile(TempF);     // Close the Temporary File
  Reset(TempF);         // Open the Temporary File for Reading

  // Copy Data from Temporary File to Output File
  Line := 0;
  while not Eof(TempF) do
  begin
    Inc(Line);                                  // Increment Line Counter

    // Get Line from Temporary File
    if Line = cTuplesCountLine then             // If it's the line we need to replace, then replace it
    begin
      // Calculate Number of Tuples Exported
      if ShowColHeaders in FOwner.FOptions then // If we're Exporting Col Headers,
        nTuple := FNumEntriesExported + 1       // then we need to +1
      else nTuple := FNumEntriesExported;       // otherwise it's just the number of entries


      // Write our Tuple Count line
      S := '0,' + IntToStr(nTuple);// This is the whole reason we needed the Temporary File.
                                   // We didn't know how many entries we were exporting when we first wrote this line down.
    end
    else ReadLn(TempF, S);         // Read line from Temporary File

    // Write Line to Output File
    WriteLn(F, S);
  end;

  CloseFile(TempF);     // Close the Temporary File
  DeleteFile(FTempDIF); // Delete Temporary File
end;

// *** TSYLKExport Methods ***

procedure TSYLKExport.WriteRow(var Items: TStringList);
const
  cSYLK_Row    = 'C;X%d;Y%d;K%s';      // X, Y, Value (Strings are quoted ", numbers are not)
  cSYLK_Format = 'F;P%d;FG0G;X%d;Y%d'; // Must go on line before actual value to be formatted but using the same cell (X,Y) coordinates
                                       // First parameter is the index of format template
var
  ColType: TColType;

  S: String;
  i, x, y: Integer;
begin
  inherited;

  with FOwner do
    for i := 0 to Items.Count - 1 do // Write Each Field
    begin
      S := Strip(Items[i]);

      ColType := GetColType(i, S); // Get Column Type

      // Calculate Cell Coordinates
      x := i + 1;

      // Calculate Actual Row Number
      if ShowColHeaders in FOptions then // If Col Headers will/are/have been shown,
        y := FNumEntriesExported + 1     // Add 1
      else y := FNumEntriesExported;     

      case ColType of
        ctNumber  :
          if not IsBlankString(S) then
            WriteLn(F, Format(cSYLK_Row, [x, y, Decimalize(GetRoundedNumber(S, TExportItemInfo(Items.Objects[i])))])); // Write as numeric data...

        ctCurrency:
          if not IsBlankString(S) then
          begin
            WriteLn(F, Format(cSYLK_Format, [cSYLK_CurrencyTemplate, x, y])); // Write Currency Formatting
            WriteLn(F, Format(cSYLK_Row, [x, y, Decimalize(S)])); // Write Currency Value
          end;

        ctDate:
          if not IsBlankString(S) then
          begin
            WriteLn(F, Format(cSYLK_Format, [cSYLK_FormatIndex_Date, x, y])); // Write Date Formatting
            WriteLn(F, Format(cSYLK_Row, [x, y, Decimalize(FloatToStr(ConvertToDate(S)))])); // Write Date Value
          end;

        ctTime:
          if not IsBlankString(S) then
          begin
            WriteLn(F, Format(cSYLK_Format, [cSYLK_FormatIndex_Time, x, y])); // Write Time Formatting
            WriteLn(F, Format(cSYLK_Row, [x, y, Decimalize(FloatToStr(StrToTime(S)))])); // Write Time Value
          end;

        ctDateTime:
          if not IsBlankString(S) then
          begin
            WriteLn(F, Format(cSYLK_Format, [cSYLK_FormatIndex_DateTime, x, y])); // Write DateTime Formatting
            WriteLn(F, Format(cSYLK_Row, [x, y, Decimalize(FloatToStr(ConvertToDateTime(S)))])); // Write DateTime Value
          end;

      else // If we don't know what it is, write it as a string (ie. ctString & ctMemo for instance would written as a string here)
        WriteLn(F, Format(cSYLK_Row, [x, y, '"' + Copy(S, 1, 255) + '"'])); // Write String Data
      end;
    end;
end;

procedure TSYLKExport.WriteHeader(var Columns: TStringList);
const
  cBeginSYLK        = 'ID;PTEXPORTX';
  cSYLKFormatRecord = 'P;P';
begin
  inherited;

  // Write DIF Header Information
  with FOwner do
  begin
    // Write Start of File Marker
    WriteLn(F, cBeginSYLK);

    // Write Format Records
    WriteLn(F, cSYLKFormatRecord + GetDateFormatString);     // Write Date Format Record [Index P0]
    WriteLn(F, cSYLKFormatRecord + GetTimeFormatString);     // Write Time Format Record [Index P1]
    WriteLn(F, cSYLKFormatRecord + GetDateTimeFormatString); // Write DateTimeFormat Record [Index P2]
    WriteLn(F, cSYLKFormatRecord + ReinforceSymbol(GetCurrencyFormatString, ';', ';')); // Write Currency Format Record [Index P3]

    // Columns Headers Row
    if ShowColHeaders in FOptions then // If we export column headers,
      WriteRow(Columns);               // Write the Column Headers Row
  end;
end;

procedure TSYLKExport.WriteEntry(var Items: TStringList; Entry: Integer);
begin
  inherited;

  WriteRow(Items); // Write the Tuple (Row)
end;

procedure TSYLKExport.WriteFooter;
const
  cSYLKColWidth = 'F;W%d %d %d'; // Column#, Column#, Col width in characters
  cEndSYLK     = 'E';
var
  i, m: Integer;
begin
  inherited;

  // Write Column Widths
  for i := 1 to FColumns.Count do
  begin
    m := PMaxColWidths^[i - 1];  // Get Max Col Width for Column
    if m > 255 then m := 255;    // Because of SYLK's 255 char max limit for strings, we need to do this
    WriteLn(F, Format(cSYLKColWidth, [i, i, m])); // Write Col Width
  end;

  // Write End of File Marker
  WriteLn(F, cEndSYLK);
end;

// *** THTMLExport Methods ***

function THTMLExport.ConvertToHTML(S: String): String;
const
  cHTML_Tab     = '&nbsp&nbsp&nbsp';
  cHTML_NewLine = cHTML_LineBreak + CRLF;
begin
  if htIgnoreLineBreaks in FOwner.FHTML_Options then        // If ignoring all linebreaks,
    Result := StripConvert(S, '', CRLF)                     // Strip out all control characters except CRLF
  else Result := StripConvert(S, cHTML_Tab, cHTML_NewLine); // otherwise, Strip all control characters and replace tabs and line feeds with HMTL equivalents
end;

function THTMLExport.LinkURLs(S: String): String;
// If S contains a URL, email, etc... it will link it

  function LinkIt(URL_Index: Integer; var S1: String): String;
  // LinkIt returns the linked part and everything before,
  // but truncates S so as to start at the position after what it has returned

    function RemoveMailTo(S2: String): String;
    // Removes the MailTo: from e-mail addresses
    const
      cMailTo          = 'mailto:';
    var
      i: Integer;
    begin
      i := Pos(cMailTo, LowerCase(S2)); // Get position of mailto:
      if i > 0 then // If found,
        Result := Copy(S2, i + Length(cMailTo), Length(S2)) // return stripped string,
      else Result := S2; // otherwise return original
    end;

  const
    cEndOfURLMarkers    = [Chr(0)..Chr(32), Chr(255)];
    cValidURLCharacters = ['a'..'z', 'A'..'Z', '?', '=', '%', '0'..'9', '@'];
    cLink               = '<a href="%s">%s</a>';
  var
    URL, URL_PostFix: String;
    IsEmailAddress: Boolean;
    i, j: Integer;
  begin
    // First find the end of the URL "word"
    for i := URL_Index to Length(S1) do
      if S1[i] in cEndOfURLMarkers then
        break;

    // Now trim the end of the URL "word" to get the index of the last URL character
    for j := URL_Index to i do
      if not (S1[j] in cValidURLCharacters) then
        break;

    // Now check if it's a potential email address
    IsEmailAddress := False; // assume it's not an e-mail address
    if S1[URL_Index] = '@' then // If potential e-mail address,
    begin
      // Find start of "word" containing '@'
      for j := URL_Index downto 1 do
        if S1[j] in cEndOfURLMarkers then
          break;

      if j > 1 then Inc(j); // adjust for cEndOfURLMarker that was found, unless there wasn't one found
      URL := Copy(S1, j, i - j); // get "word"

      URL_PostFix := Copy(URL, Pos('@', URL) + 1, Length(URL)); // Get all stuff after '@' symbol -> URL_PostFix
      if (Pos('.', URL_PostFix) > 1) and (Pos('@', URL) > 1) then // if '.' appears after '@' symbol and
      begin                                           // there are characters before the '@' symbol,
        IsEmailAddress := True; // we assume it's an e-mail address
        URL_Index := j;
      end
      else
      begin // Not an e-mail address but had the '@' symbol, so lets quit because it's not a URL
        Result := Copy(S1, 1, URL_Index);          // Return stuff before '@' symbol
        S1 := Copy(S1, URL_Index + 1, Length(S1)); // Skip processing to after '@' symbol
        Exit;
      end;
    end;

    // If E-mail address then
    if IsEmailAddress then
      URL := 'mailto:' + URL // add 'mailto:'
    else URL := Copy(S1, URL_Index, i - URL_Index); // else get the URL

    // Return the entire string up to and including the URL
    Result := Copy(S1, 1, URL_Index - 1) +
              Format(cLink, [URL, RemoveMailTo(URL)]); // Now we Link the URL

    // Truncate S so as not to include what we just returned
    S1 := Copy(S1, i, Length(S1));
  end;

type
  TURL_Constants = array [1..7] of String;
const
  cURLs: TURL_Constants =
   ('http://',
    'https://',
    'ftp://',
    'news://',
    'nntp://',
    'mailto:',
    '@');
var
  i, u, URL_Pos: Integer;
begin
  // If not autolinking URLs then quit
  if not (htAutoLink in FOwner.FHTML_Options) then
  begin
    Result := S;
    Exit;
  end;

  // Initialize
  Result := '';

  // Link All URLs in String
  repeat
    URL_Pos := 0;
    for i := Low(TURL_Constants) to High(TURL_Constants) do // Loop through all possible URLs
    begin
      // Find start of URL
      u := Pos(cURLs[i], LowerCase(S));

      // If it's before the last URL, then flag this as the next one to process
      if (u > 0) and ((URL_Pos = 0) or (u < URL_Pos)) then
        URL_Pos := u;
    end;

    // if URL found, then link it
    if URL_Pos > 0 then
      Result := Result + LinkIt(URL_Pos, S);

  until URL_Pos = 0;

  if Result = '' then        // If No URLs were found,
    Result := S              // return original string
  else Result := Result + S; // otherwise, add the last bit of the original text to the result
end;

function THTMLExport.EnsureNotEmpty(S: String): String;
begin
  Result := Trim(S);         // Trim String
  if Length(Result) = 0 then // If Empty,
    Result := cEmptyCell;    // put Special "Empty Cell" Tags so the browser export looks good with gridlines
end;

function THTMLExport.GetElement(i: Integer): String;
begin
  Result := GetToken(FSelectedTemplate, i); // Get the HTML Element at the 'i' position
end;

procedure THTMLExport.WriteHeader(var Columns: TStringList);
type
  TBoldTags = record
    Tag,
    EndTag: String;
  end;
var
  Bold: TBoldTags;
  S: String;
  i, BorderSize: Integer;
begin
  inherited;

  with FOwner do
  begin
    // Get Selected HTML Template
    if IsBlankString(FHTML_CustomTemplate) then            // If HTML_CustomTemplate is Blank,
      FSelectedTemplate := cHTML_Templates[FHTML_Template] // use selected pre-built template
    else FSelectedTemplate := FHTML_CustomTemplate;        // else use user specified CustomTemplate

    // Write HTML Header
    WriteLn(F, Format(cBeginHTML, [ConvertToHTML(FTitle),
                                   GetElement(cHTML_BGColor),
                                   GetElement(cHTML_Link),
                                   GetElement(cHTML_VLink),
                                   GetElement(cHTML_TableFontColor),
                                   GetElement(cHTML_ALink),
                                   GetElement(cHTML_Text),
                                   GetElement(cHTML_TextFont)]));

    // Write Title
    if htDisplayTitle in FHTML_Options then
      WriteLn(F, Format(cHTML_Title, [ConvertToHTML(FTitle)]));

    // Write Time Stamp
    if TimeStamp in FOptions then
      WriteLn(F, GetTimeStamp + cHTML_LineBreak_x2);

    // Write User Defined Header
    if Has(FHeader.Text) then
    begin
      S := LinkURLs(ConvertToHTML(FHeader.Text));
      if not (htIgnoreLineBreaks in FHTML_Options) then  // Add Line Break (unless we are told not to)
        if htHorzRules in FHTML_Options then // If htHorzRules set,
          S := S + cHTML_HorzRule            // Add Horizontal Rule (Separator)
        else S := S + cHTML_LineBreak;       // otherwise, add a simple line break

      WriteLn(F, S);
    end;

    // Write Begin Table Code
    if htShowGridLines in FHTML_Options then
      BorderSize := 1
    else BorderSize := 0;
    WriteLn(F, Format(cBeginTable, [BorderSize, GetElement(cHTML_TableBGColor)]));

    // Begin Row
    if ShowColHeaders in FOptions then
    begin
      WriteLn(F, Format(cBeginColoredRow, [GetElement(cHTML_ColHeadersBGColor)]));

      // Write Column Headers
      for i := 0 to Columns.Count - 1 do
      begin
        // Determine if Column Headers Should be Bold
        if htColHeadersBold in FHTML_Options then
        begin
          Bold.Tag    := cBeginBold;
          Bold.EndTag := cEndBold;
        end;

        // Write Column Header
        Writeln(F, Format(cBeginTitleCell, [Bold.Tag,
                                            GetElement(cHTML_ColHeadersFontColor)]) +
                   EnsureNotEmpty(ConvertToHTML(Columns[i])) +
                   Format(cEndTitleCell, [Bold.EndTag]));
      end;

      // End Row
      WriteLn(F, cEndRow);
    end;
  end;
end;

procedure THTMLExport.WriteEntry(var Items: TStringList; Entry: Integer);
var
  i: Integer;
begin
  inherited;

  // Begin Row
  if htOddRowColoring in FOwner.FHTML_Options then // If Odd Row Coloring Set,
  begin
    if Odd(Entry) then // If Odd Row,
      WriteLn(F, Format(cBeginColoredRow, [GetElement(cHTML_OddRowBGColor)])) // Write Odd Row BGColor
    else WriteLn(F, Format(cBeginColoredRow, [GetElement(cHTML_TableBGColor)])); // else write Even Row BGColor
  end
  else
    WriteLn(F, cBeginRow); // Otherwise, write Regular Row

  // Write the Cell
  with FOwner do
    for i := 0 to Items.Count - 1 do
      WriteLn(F, cBeginCell + EnsureNotEmpty(LinkURLs(ConvertToHTML(GetDisplayString(Items, i)))) + cEndCell);

  // End Row
  WriteLn(F, cEndRow);
end;

procedure THTMLExport.WriteFooter;
var
  S: String;
begin
  inherited;

  // Write End Table Tag
  Write(F, cEndTable);

  // Write HTML Footer
  with FOwner do
    if Has(FFooter.Text) then
    begin
      S := LinkURLs(ConvertToHTML(FFormatFooter.Text));
      if not (htIgnoreLineBreaks in FHTML_Options) then // Add Line Break (unless we are told not to)
        if htHorzRules in FHTML_Options then // If htHorzRules set,
          S := cHTML_HorzRule + S            // Add Horizontal Rule (Separator)
        else S := cHTML_LineBreak + S;       // otherwise, add a simple line break

      WriteLn(F, S);
    end;

  // Write End HTML Tag
  Write(F, cEndHTML);
end;

// *** TRichTextExport Methods ***

constructor TRichTextExport.Create(aOwner: TExportXComponent; aExportFile: String; aExportType: TExportType);
begin
  inherited;

  // Adjust Default Col Width Spacing for RTF Exports
  FOwner.FDefaultColWidthSpacing := cRTF_DefaultColWidthSpacing;
end;

function TRichTextExport.ConvertToRTF(S: String): String;
  function ReinforceRTF(S: String): String;
  begin
    Result := ReinforceSymbol(S, '\', '\');
    Result := ReinforceSymbol(Result, '{', '\');
    Result := ReinforceSymbol(Result, '}', '\');
  end;

  function _7bit(S: String): String; // Converts to 7bit characters
  var
    i: Integer;
  begin
    Result := '';
    for i := 1 to Length(S) do
      if Ord(S[i]) > 127 then                                        // If 8bit Character,
        Result := Result + '\''' + LowerCase(IntToHex(Ord(S[i]), 2)) // make 7bit ie. \'ac
      else Result := Result + S[i];                                  // Otherwise, just pass-through
  end;

const
  cRTF_Tab     = '\tab ';
  cRTF_NewLine = '\par' + CRLF;
begin
  Result := StripConvert(_7bit(ReinforceRTF(S)), cRTF_Tab, cRTF_NewLine);
end;

procedure TRichTextExport.WriteColFormatInfo(Cell, Prefix, Suffix, Regular: String);
const
  cMaxCharWidth = 90; // This is the character space we want to allocate for each character
                      // in the column heading to properly space out the columns.
var
  i, x, ColWidth: Integer;
  Template: String;
begin
  // Initialize
  ColWidth := 0;

  // Format Columns
  with FColumns do
  for i := 0 to Count - 1 do
  begin
    // Get Colummn Format Template String
    if Count = 1     then Template := Prefix + Cell + Suffix else // If only one Col
    if i = 0         then Template := Prefix + Cell          else // If First Col
    if i = Count - 1 then Template := Cell + Suffix          else // If Last Col
      Template := Regular; // If Regular Col


    // Calculate the "x"-factor
    if (i = 0) and (NumberRows in FOwner.FOptions) then // If 'Item #' then
      x := 6 // Needs to be done because GetColWidth would return 8 or 9 or something for 'Item #' - since padded-text exports truncate instead of wrap-around values
    else x := GetColWidth(i) + 3;                       // else, if data column.

    // Calculate Column Width
    ColWidth := ColWidth + x * cMaxCharWidth;

    // Write the Format Line to File
    WriteLn(F, Format(Template, [ColWidth]));
  end;
end;

procedure TRichTextExport.WriteHeader(var Columns: TStringList);
var
  i: Integer;
begin
  inherited;

  // Rich Text Header & Column Header Stuff
  with FOwner do
  begin
    // Write RichText Header
    with FRTF_Options do
      WriteLn(F, Format(cRTF_Header, [LeftMargin,      // Left Margin
                                      RightMargin,     // Right Margin
                                      TopMargin,       // Top Margin
                                      BottomMargin])); // Bottom Margin

    // Write Title
    if Has(FTitle) then
      WriteLn(F, Format(cRTF_Title, [ConvertToRTF(FTitle)]));

    // Begin Text
    WriteLn(F, cRTF_BeginText);

    // Write TimeStamp
    if TimeStamp in FOptions then
      WriteLn(F, Format(cRTF_TimeStamp, [GetTimeStamp]));

    // Write User Defined Header
    if Has(FHeader.Text) then
      WriteLn(F, Format(cRTF_UserHeader, [ConvertToRTF(FHeader.Text)]));

    // Write Column Headers
    if ShowColHeaders in FOptions then
    begin
      // Write Before Col Heading Formatting Info
      WriteColFormatInfo(cRTF_ColFormatHeadingCell,
                         cRTF_ColFormatHeadingPrefix,
                         cRTF_ColFormatHeadingSuffix,
                         cRTF_ColFormatHeading);

      // Write Column Headers
      for i := 0 to Columns.Count - 1 do
        Writeln(F, ConvertToRTF(Columns[i]) + cRTF_AfterCell);
    end;
  end;
end;

procedure TRichTextExport.WriteEntry(var Items: TStringList; Entry: Integer);
var
  i, Padding: Integer;
begin
  inherited;

  // Write Before Row Formatting Info
  if Entry = 1 then
      WriteColFormatInfo(cRTF_ColFormatInitialRowCell,
                       cRTF_ColFormatInitialRowPrefix,
                       cRTF_ColFormatInitialRowSuffix,
                       cRTF_ColFormatInitialRowCell) else // First Row

  if not FOwner.MoreExportItems then
    WriteColFormatInfo(cRTF_ColFormatFinalRowCell,
                       cRTF_ColFormatFinalRowPrefix,
                       cRTF_ColFormatFinalRowSuffix,
                       cRTF_ColFormatFinalRow) else // Last Row

    WriteLn(F, cRTF_BeforeRow); // Regular Row

  // Number of Empty Cells at End of Row
  Padding := FColumns.Count - Items.Count;

  // Write the Cells with Data
  for i := 0 to Items.Count - 1 do
    WriteLn(F, ConvertToRTF(GetDisplayString(Items, i)) + cRTF_AfterCell);

  // Write the Padding (Empty Cells at End of Row)
  for i := 1 to Padding do
    WriteLn(F, cRTF_AfterCell)
end;

procedure TRichTextExport.WriteFooter;
begin
  inherited;

  // Write RichText Footer
  WriteLn(F, Format(cRTF_Footer, [ConvertToRTF(FOwner.FFormatFooter.Text)]));
end;

// *** TCustomPrint Methods ***

constructor TCustomPrintExport.Create(aOwner: TExportXComponent; aExportFile: String; aExportType: TExportType);
begin
  inherited;

  FCreatedSuccessfully := False;

  // Show Print Dialog
  with TPrintDialog.Create(nil) do
  try
    FUserCancelled := not Execute;

  finally
    Free;
  end;

  if not FUserCancelled then
  begin
    if Assigned(FProgForm) then
      FProgForm.Caption := cPrintingCaption;

    // Assign Printer
    Printer.Title := FOwner.FTitle; // Set Print Job Document Name
    AssignPrn(PrintText);
    Rewrite(PrintText);
  end
  else // If User Cancelled, get rid of Progress Form
    if Assigned(FProgForm) then
    begin
      FProgForm.Free;
      FProgForm := nil;
    end;

  FCreatedSuccessfully := True;
end;

procedure TCustomPrintExport.WriteHeader(var Columns: TStringList);
begin
  inherited;

  // Set Title Font
  with Printer.Canvas.Font do
  begin
    Name  := cDetailedTitleName;
    Size  := cDetailedTitleSize;
    Style := cDetailedTitleStyle;
  end;
end;

destructor TCustomPrintExport.Destroy;
begin
  if FCreatedSuccessfully and not FUserCancelled then
    CloseFile(PrintText);

  inherited;
end;

procedure TCustomPrintExport.HandleError(Exception: TObject);
begin
  MessageDlg(cPrintError, mtError, [mbOk], 0);
end;

// *** TPrintExport Methods ***

procedure TPrintExport.WriteHeader(var Columns: TStringList);
begin
  // Set Font
  with Printer.Canvas.Font do
  begin
    Name  := cDefaultFontName;
    Size  := cDefaultFontSize;
    Style := cDefaultFontStyle;
  end;

  Write(PrintText, GetPaddedHeader(Columns));
end;

procedure TPrintExport.WriteEntry(var Items: TStringList; Entry: Integer);
begin
  Write(PrintText, GetPaddedEntry(Items, Entry));
end;

procedure TPrintExport.WriteFooter;
var
  S: String;
begin
  inherited;

  // Make Italic
  Printer.Canvas.Font.Style := [fsItalic];

  S := GetPaddedFooter;
  if Length(S) > 0 then
    Write(PrintText, S);
end;

// *** TPrintDetails Methods ***

procedure TDetailedPrintExport.WriteHeader(var Columns: TStringList);
begin
  inherited;

  CalculateMaxColumnLength;  

  with FOwner do
  begin
    if Has(FTitle) then
    begin
      WriteLn(PrintText, FTitle);
      WriteLn(PrintText, '');
    end;

    if Has(FHeader.Text) then
    begin
      // Set Font
      with Printer.Canvas.Font do
      begin
        Name  := cDetailedHeaderName;
        Size  := cDetailedHeaderSize;
        Style := cDetailedHeaderStyle;
      end;

      WriteLn(PrintText, FHeader.Text);
      WriteLn(PrintText, '');
    end;
  end;

  // Set Font for Printed Entries
  with Printer.Canvas.Font do
  begin
    Name  := cDetailedFontName;
    Size  := cDetailedFontSize;
    Style := cDetailedFontStyle;
  end;
end;

procedure TDetailedPrintExport.WriteEntry(var Items: TStringList; Entry: Integer);
var
  i: Integer;
  S: String;
begin
  inherited;

  with FOwner do
    for i := 0 to Items.Count - 1 do // Write Each Field
    begin
      // Get Col Header
      if ShowColHeaders in FOptions then
      begin
        if Has(FColumns[i]) then
          S := Pad(FColumns[i], FMaxColumnLength, False)  + ': '
        else S := '';
      end
      else S := '';

      WriteLn(PrintText, S + Items[i]);
    end;

  // Write a Spacing Line (as long as it's not the last line)
  if FOwner.MoreExportItems then
    WriteLn(PrintText, '');
end;

procedure TDetailedPrintExport.WriteFooter;
begin
  inherited;

  with FOwner do
    if Has(FFormatFooter.Text) then
    begin
      // Set Font
      with Printer.Canvas.Font do
      begin
        Name  := cDetailedFooterName;
        Size  := cDetailedFooterSize;
        Style := cDetailedFooterStyle;
      end;

      WriteLn(PrintText, '');
      WriteLn(PrintText, '-----');

      // Write Footer
      WriteLn(PrintText, FFormatFooter.Text);
    end;
end;

// *** Misc. Functions ***

{$HINTS OFF}
function IsCardinal(S: String): Boolean;
var
  TestInt: Integer;
  ErrorCode: Integer;
begin
  Val(S, TestInt, ErrorCode); // Try to Convert to Integer Point Value
  Result := (ErrorCode = 0) and (StrToInt(S) >= 0);
end;
{$HINTS ON}

function IsBlankString(S: String): Boolean;
begin
  Result := Length(Trim(S)) = 0;
end;

function GetNumDataColumns(ColCount: Integer; aOptions: TExportOptions): Integer;
begin
  Result := ColCount;
  if NumberRows in aOptions then Dec(Result);
end;

function GetUnderline(aLength: Integer): String;
var
  i: Integer;
begin
  SetLength(Result, aLength);
  for i := 1 to aLength do
    Result[i] := cUnderlineChar;
end;

function GetSingleColEntry(aEntry: String; aEntryIndex: Integer; aOptions: TExportOptions): String;
begin
  Result := '';
  if NumberRows in aOptions then
    Result := IntToStr(aEntryIndex) + cNumberPostFix + ' ';
  Result := Result + aEntry ;
end;

function GetToken(S: String; i: Integer): String;
// ie. a;b;c has 3 tokens. a = 1, b = 2, c =3
const
  cInvalidIndexError = 'There is no Token "%d" in the String "%s"';
  cTokenSeparator    = ';';
var
  Old_S: String;
  p, Old_i: Integer;
begin
  // Token index must be greater than 0
  if i < 1 then
    raise Exception.CreateFmt(cInvalidIndexError, [i, S]);

  // Grab Token
  Old_i := i; // Save for error message (just in case)
  Old_S := S; // ...
  while i > 0 do
  begin
    p := Pos(cTokenSeparator, S); // Get Next Token
    if p > 0 then
    begin
      Result := Copy(S, 1, p - 1);        // Get Token
      S := Copy(S, p + 1, Length(S)); // Copy stuff after Token
    end
    else
      if i > 1 then // If it's not the last token, it must be an index which is too high
        raise Exception.CreateFmt(cInvalidIndexError, [Old_i, Old_S])
      else Result := S; // If no more tokens then we've got the last token

    Dec(i);
  end;
end;

function ExpandTempDir(TheFile: String): String;
// Returns the Temp Directory, if that fails, then the windows dir
var
   Buff: array[1..cMaxPathSize] of Char;
begin
  // Get User Specified Temp Dir
  Result := FUserTempDir;

  // If no user Temp Dir is specified, calculate it ourselves
  if IsBlankString(FUserTempDir) then
  begin
    if GetTempPath(cMaxPathSize, @Buff) = 0 then // if failed
     if GetWindowsDirectory(@Buff, cMaxPathSize) = 0 then // if still failed
        Result := 'c:\'; {then hope for the best}

    if Length(Result) = 0 then // if a path was returned, convert it
      Result := pchar(@Buff);
  end;

  // Attach File name to Temp Dir path
  if Result[Length(Result)] <> '\' then Result := Result + '\';
    Result := Result + TheFile;
end;

function ReinforceSymbol(S: String; Symbol, Reinforcer: Char): String;
// If we want to reinforce '"' with another '"' then ReinforceSymbol('Hello "Joe"', '"', '"')
// would yield 'Hello ""Joe""'. ReinforceSymbol('{FONT}', '{', '\') would yield
// '\{FONT}'
var
  i: Integer;
begin
  i := Pos(Symbol, S); // Find position of Symbol

  if i = 0 then
    Result := S
  else
  begin
    Result := '';
    repeat
      Result := Result + Copy(S, 1, i - 1) + Reinforcer + S[i]; // Reinforce
      S := Copy(S, i + 1, Length(S));
      i := Pos(Symbol, S); // Find position of quote symbol
    until i = 0;
    Result := Result + S;
  end;
end;

function Has(S: String): Boolean;
begin
  Result := Length(Trim(S)) > 0
end;

function ClassExists(aClass: String): Boolean;
begin
  with TRegistry.Create do
  try
    RootKey := HKEY_CLASSES_ROOT;
    Result  := KeyExists(aClass);

  finally
    Free;
  end;
end;

function GetExportType(Choice: String): TExportType;
var
  i: TExportType;
begin
  Result := xText; // Default to Text

  for i := Low(cExportTypes) to High(cExportTypes) do
    if Choice = cExportTypes[i] then
    begin
      Result := TExportType(i);
      break;
    end;
end;

initialization

  // Create TempFiles Variable
  TempFiles := TStringList.Create;

finalization

  // Remove All Temporary Files & Free Temp Files Variable
  if Assigned(TempFiles) then
  try
    while TempFiles.Count > 0 do
    begin                                                                       
      DeleteFile(TempFiles[0]);
      TempFiles.Delete(0);
    end;

  finally
    TempFiles.Free;
  end;

  // Find All old Temporary Files and remove them
  if FindFirst(ExpandTempDir(cTempFilePrefix + '*.*'), faAnyFile, SearchRec) = 0 then
  try
    repeat
      DeleteFile(ExpandTempDir(SearchRec.Name));
    until FindNext(SearchRec) <> 0;

  finally
    FindClose(SearchRec);
  end;
end.
