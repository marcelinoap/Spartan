{------------------------------------------------------------------------------}
{ TExportX Components Localization Unit - English Language                     }
{------------------------------------------------------------------------------}
{ Before Translating this file yourself, please ensure that there is not a     }
{ translation already available in your language on our Web Site.              }
{ http://www.igather.com/components                                            }
{                                                                              }
{ To localize all Y-Tech Export Components, just change these String Constants }
{ to phrases in your own language. Please make sure you don't change the name  }
{ of constants, or change them in any other way - you must leave all %s, etc...}
{ markers in the string. Only change the actual words itself. If you do        }
{ anything else, TExportListView might not compile or run properly.            }
{                                                                              }
{ You have 2 options: 1) (D3+ Only) you can change the resourcestrings after   }
{ the unit has been compiled or 2) you can change the strings right in here.   }
{ Note: #2 is the only option for D2 Users.                                    }
{                                                                              }
{ Please send any decent translations you have made to ycomp@hotpop.com so we  }
{ can include them on our web site for the benefit of other people that speak  }
{ your language.                                                               }   
{------------------------------------------------------------------------------}

unit EC_Strings;

interface

{$I CompConditionals.inc}

{$IFDEF Delphi3Up}
ResourceString
{$ELSE}
const
{$ENDIF}
  // These are the names that will appear in the Choose Dialog, you'll want to translate these
  // for sure.
  ccHTML           = 'HTML';
  ccMicrosoftWord  = 'Microsoft Word';
  ccMicrosoftExcel = 'Microsoft Excel';
  ccText           = 'Texto';
  ccRichText       = 'Rich Text';
  ccCSVText        = 'Comma-Delimited Text (CSV)';
  ccTabText        = 'Tab-Delimited Text';
  ccDif            = 'Data Interchange Format (DIF)';
  ccSYLK           = 'SYLK Format';
  ccClipboard      = 'Clipboard';

  cFilesFilter       = 'Arquivos %s (*.%s)|*.%s|Todos os Arquivos (*.*)|*.*';
  cExportTitle       = 'Exportar para o Arquivo %s';

  // "Choose" Dialog Strings
  cChooseDialogCaption   = 'Exportar %d Itens';
  cChooseDialogCaption2  = 'Exporta Itens';
  cChooseDialogStatusMsg = '%d Itens Exportados em %.2f Segundos.';
  cChooseOkCaption       = '&Ok';
  cChooseCancelCaption   = '&Cancelar';
  cChooseInstructions    = 'Por favor, escolha um Formato para Exportar';
  cChooseScreenCaption   = '&Ver na Tela';
  cChooseFileCaption     = 'Exportar para &Arquivo';

  // Misc.
  cExcelRightFooter = 'Página &P de &N'; // Don't touch &P or &N! Only change word for 'page' and 'of'
  cNumberRowsCaption = 'Item'; // Row Numbering Column Header

  // Progress Constatns
  cPrintingCaption = 'Imprimindo...';
  cExportingCaption = 'Exportando...';

  // Misc. Messages
  cCreatedMsg         = 'Gerado:';
  cClipSuccessfulMsg  = 'Exportado corretamente para o Clipboard';
  cExcelPleaseWaitMsg = 'Por favor Aguarde... Exportação para Excel em Progresso.';

  // Error Messages
  cGenericExportError  = 'Exportação Falhou.';
  cTextFileError       = 'Export failed. Try again with a different file name. Please ensure that ' +
                         'there is enough room on the destination drive and that the target drive is ' +
                         'not write-protected.' + #13#10#13#10 +
                         'Network users: Please ensure that the network ' +
                         'drive/folder you are trying to write to is connected and accessible.';
  cPrintError          = 'Could not complete Print operation successfully.';
  cDefShellViewError   = 'Error occured while trying to view data exported to the %s format.' + #13#10#13#10 +
                         'If this problem persists, you may want to consider re-installing the %s ' +
                         'Viewer Application.';

  cNoDataError       = 'Não há dados para Exportar.';
  cNoneSelectedError = 'Pelo menos um ítem deve ser selecionado para Exportar.';
  cNoColumnsError    = #13#10#13#10 + 'Por favor cheque se a culuna está visível e possui valor.';

  // OLE Error Messages - OLE is not used anymore, but this is left in, just in case someone is using v2.70 or earlier
  cGenericOLEError1    = 'Não há comunicação com ';
  cGenericOLEError2    = '.' + #13#10 + 'Tente novamente e se ainda falhas então '+
                         'reinicie o computador.';

const
  // Error Messages that only the Developer should see, no need to translate these.
  cNoComponentError     = 'Você deve associar um %s antes de Exportar/Imprimir.';
  cNotActiveError       = 'Dataset associado não ativo. Ative a antes de ' +
                          'tentaar Exportar/Imprimir.';
  cNoDataSource         = '%s não possui DataSource associado';
  cNoDataSet            = '%s has possui DataSet associado';

implementation

end.
