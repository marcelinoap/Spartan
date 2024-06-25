{------------------------------------------------------------------------------}
{ TExportX-DB Suite v3.01 (c) 1998-99 Y-Tech Corporation      April 25th, 1999 }
{------------------------------------------------------------------------------}
{ TExportX-DB components can Print or Export any DBGrid, Table, Query or other }
{ DataSet to 10 Formats (incl. HTML, Microsoft Word & Excel) Exports to Screen }
{ or File. Can Export if target App not installed. Shows Progress. Very Fast.  }
{ (Source w/Registration)                                                      }
{                                                                              }
{ Please visit our Web Site: http://www.igather.com/components to try out the  }
{ many other freeware & shareware components available there.                  }
{------------------------------------------------------------------------------}

{------------------------------------------------------------------------------}
{ Components Included                                                          }
{------------------------------------------------------------------------------}
{ TExportDBGrid : Export any DBGrid                                            }
{ TExportTable  : Export any Table, Query or other DataSet                     }
{                                                                              }
{ Registered Version also includes:                                            }
{ - TExportListView    : Export any TListView or descendant.                   }
{ - TExportEnhListView : Export Brad Stower's TEnhListView.                    }
{ - TExportExtListView : Export Brad Stower's TExtListView.                    }
{ - TExportStringGrid  : Export any TStringGrid.                               }
{ - TExportStrings     : Export any TStrings or TStringList.                   }
{------------------------------------------------------------------------------}

{------------------------------------------------------------------------------}
{ Installation                                                                 }
{------------------------------------------------------------------------------}
{ First UnZip the files. You *MUST* make sure that WinZip (or whatever program }
{ you use) creates the subdirectories and does not unzip everything into one   }
{ folder. (ie. there should be Help, Demo and other subdirectories)            }
{                                                                              }
{ 1) Please register ExportXDBReg.pas                                          }
{                                                                              }
{ Please ensure you remove the unregistered version before installing the      }
{ registered version.                                                          }
{                                                                              }
{ Important! Since many Y-Tech Components (ie. those made by Y-Tech Corp.)     }
{ share the same units, you must put all code files (.dcu, .pas, .dfm,         }
{ .dcr, etc...) for all Y-Tech Components into the same directory. Otherwise   }
{ Delphi will not compile the library/package properly.                        }
{                                                                              }
{ Registered Users: Installing the bonus TExportX Non-DB Components            }
{ -----------------------------------------------------------------            }
{ 1) Please register ExportXReg.pas                                            }
{ 2) For Brad Stower's TEnh/TExtListView support, register ExportExtLVReg.pas  }
{ 3) For TExportStrings, please register ExportStringsReg.pas. This is a       }
{    Freeware component, so it will work even if you don't decide to purchase  }
{    The TExportX Component Suite.                                             }
{ 4) For TDebugLog users, please refer to TDebugLog's "install.txt" after      }
{    installation of this component.                                           }
{                                                                              }
{ If you have problems installing the components, please read the              }
{ "Troubleshooting.txt" file.                                                  }
{------------------------------------------------------------------------------}

{------------------------------------------------------------------------------}
{ Questions & Comments                                                         }
{------------------------------------------------------------------------------}
{ All questions and comments should be sent to ycomp@hotpop.com                }
{------------------------------------------------------------------------------}

{------------------------------------------------------------------------------}
{ Usage                                                                        }
{------------------------------------------------------------------------------}
{ - Always read "readme.txt" (you may be reading it now)                       }
{ - Upgraders should read "upgraders.txt" and "history.txt" for important      }
{   information.                                                               }
{ - Please look at "info.htm"                                                  }
{ - Please refer to the help file.                                             }
{------------------------------------------------------------------------------}

{------------------------------------------------------------------------------}
{ License Agreement & Disclaimer                                               }
{------------------------------------------------------------------------------}
{ By use of this component, you have agreed to the following:                  }
{                                                                              }
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

{------------------------------------------------------------------------------}
{ International Language Support                                               }
{------------------------------------------------------------------------------}
{ There are 2 ways of implementing support for your own language:              }
{                                                                              }
{ 1. Get the EC_Strings.pas file for your language and follow the              }
{    installation instructions that come with it. Try downloading it from the  }
{    Y-Tech Components web site (http://www.igather.com/components).           }
{                                                                              }
{ 2. Translate it yourself. (if there is no EC_Strings.pas for your language,  }
{    you must do it this way). It's simple really, just find the               }
{    EC_Strings.pas unit that came with TExportListView and translate the      }
{    strings inside. There's not that many, so it's not a big job. Then        }
{    recompile the package/component library that your component is installed  }
{    in. Finally, if you think you made a half-decent translation, please      }
{    e-mail it to Y-Tech at ycomp@hotpop.com                                   }
{------------------------------------------------------------------------------}

{------------------------------------------------------------------------------}
{ File Format Notes                                                            }
{------------------------------------------------------------------------------}
{ Remember... If you have problems with any Export format, you can always      }
{ disable it using the "AllowedTypes" property.                                }
{                                                                              }
{ XLS  + Microsoft Excel has a 255 char Limit for Strings and Memos.           }
{                                                                              }
{ CSV  + This format is slightly less universally compatible than Tab-Text,    }
{        DIF or SYLK since it is (believe it or not!) not always               }
{        comma-separated. It uses the Control Panel/Regional/Number/List       }
{        Separator value to determine the delimiter. Usually this is a ',' but }
{        in some countries like Germany it will be a semi-colon as ',' is used }
{        for currencies. This is is keeping with Microsoft Excel's practices.  }
{                                                                              }
{ SYLK + This format has the advantage over DIF, Tab-Text and CSV in that it   }
{        maintains the widths of the columns, so imports into programs like    }
{        Excel will look nicer. But it has one minor disadvantage:             }
{                                                                              }
{        - SYLK has a 255 char Limit for Strings and Memos.                    }
{------------------------------------------------------------------------------}

unit ExportXDBReg;

interface

uses EC_DBGrid, EC_Table;

const
  {$I CompConstants.inc} // Component Constants & Conditional Defines

procedure Register;

implementation

uses Classes;

procedure Register;
begin
  RegisterComponents(cExportTab, [TExportDBGrid, TExportTable]);
end;

end.
