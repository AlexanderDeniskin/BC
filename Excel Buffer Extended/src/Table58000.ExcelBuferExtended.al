namespace ASD.Excel;

using System.IO;
using System.Reflection;
table 58000 "Excel Buffer Extended"
{
    Caption = 'Excel Buffer Extended';
    ReplicateData = false;

    fields
    {
        field(1; "Row No."; Integer)
        {
            Caption = 'Row No.';
            DataClassification = SystemMetadata;

            trigger OnValidate()
            begin
                xlRowID := '';
                if "Row No." <> 0 then
                    xlRowID := Format("Row No.");
            end;
        }
        field(2; xlRowID; Text[10])
        {
            Caption = 'xlRowID';
            DataClassification = SystemMetadata;
        }
        field(3; "Column No."; Integer)
        {
            Caption = 'Column No.';
            DataClassification = SystemMetadata;

            trigger OnValidate()
            var
                ExcelBufferHelper: Codeunit "Excel Buffer Helper";
            begin
                xlColID := ExcelBufferHelper.GetColumnCode("Column No.");
            end;
        }
        field(4; xlColID; Text[10])
        {
            Caption = 'xlColID';
            DataClassification = SystemMetadata;
        }
        field(5; "Cell Value as Text"; Text[250])
        {
            Caption = 'Cell Value as Text';
            DataClassification = SystemMetadata;
        }
        field(6; Comment; Text[500])
        {
            Caption = 'Comment';
            DataClassification = SystemMetadata;
        }
        field(7; Formula; Text[500])
        {
            Caption = 'Formula';
            DataClassification = SystemMetadata;
        }
        field(8; Bold; Boolean)
        {
            Caption = 'Bold';
            DataClassification = SystemMetadata;
        }
        field(9; Italic; Boolean)
        {
            Caption = 'Italic';
            DataClassification = SystemMetadata;
        }
        field(10; Underline; Boolean)
        {
            Caption = 'Underline';
            DataClassification = SystemMetadata;
        }
        field(11; NumberFormat; Text[30])
        {
            Caption = 'NumberFormat';
            DataClassification = SystemMetadata;
        }
        field(12; Formula2; Text[500])
        {
            Caption = 'Formula2';
            DataClassification = SystemMetadata;
        }
        field(13; Formula3; Text[500])
        {
            Caption = 'Formula3';
            DataClassification = SystemMetadata;
        }
        field(14; Formula4; Text[500])
        {
            Caption = 'Formula4';
            DataClassification = SystemMetadata;
        }
        field(15; "Cell Type"; Option)
        {
            Caption = 'Cell Type';
            DataClassification = SystemMetadata;
            OptionCaption = 'Number,Text,Date,Time';
            OptionMembers = Number,Text,Date,Time;
        }
        field(16; "Double Underline"; Boolean)
        {
            Caption = 'Double Underline';
            DataClassification = SystemMetadata;
        }
        field(17; "Cell Value as Blob"; Blob)
        {
            Caption = 'Cell Value as Blob';
            DataClassification = SystemMetadata;
        }
        field(50000; "Font Name"; Text[30])
        {
            Caption = 'Font Name';
            DataClassification = ToBeClassified;
        }
        field(50001; "Font Size"; Decimal)
        {
            Caption = 'Font Size';
            DataClassification = ToBeClassified;
        }
        field(50002; "Font Color"; Text[30])
        {
            Caption = 'Font Color';
            DataClassification = ToBeClassified;
        }
        field(50003; Strikethrough; Boolean)
        {
            Caption = 'Strikethrough';
            DataClassification = ToBeClassified;
        }
        field(50004; "VertAlign Effect"; Option)
        {
            Caption = 'Vertical Align Effect';
            DataClassification = ToBeClassified;
            OptionCaption = ' ,Superscript,Subscript';
            OptionMembers = " ",Superscript,Subscript;
        }
        field(50010; "Background Color"; Text[30])
        {
            Caption = 'Background Color';
            DataClassification = ToBeClassified;
        }
        field(50011; "Patern Style"; Option)
        {
            Caption = 'Patern Style';
            DataClassification = ToBeClassified;
            OptionCaption = 'None,Solid,75% Gray,50% Gray,25% Gray,12.5% Gray,6.25% Gray,Horizontal Stripe,Vertical Stripe,Reverse Diagonal Stripe,Diagonal Stripe,Diagonal Crosshatch,Thick Diagonal Crosshatch,Thin Horizontal Stripe,Thin Vertical Stripe,Thin Reverse Diagonal Stripe,Thin Diagonal Stripe,Thin Horizontal Crosshatch,Thin Diagonal Crosshatch';
            OptionMembers = None,Solid,"75% Gray","50% Gray","25% Gray","12.5% Gray","6.25% Gray","Horizontal Stripe","Vertical Stripe","Reverse Diagonal Stripe","Diagonal Stripe","Diagonal Crosshatch","Thick Diagonal Crosshatch","Thin Horizontal Stripe","Thin Vertical Stripe","Thin Reverse Diagonal Stripe","Thin Diagonal Stripe","Thin Horizontal Crosshatch","Thin Diagonal Crosshatch";
        }
        field(50012; "Patern Color"; Text[30])
        {
            Caption = 'Patern Color';
            DataClassification = ToBeClassified;
        }
        field(50013; "Gradient Color 1"; Text[30])
        {
            Caption = 'Gradient Color 1';
            DataClassification = ToBeClassified;
        }
        field(50014; "Gradient Color 2"; Text[30])
        {
            Caption = 'Gradient Color 2';
            DataClassification = ToBeClassified;
        }
        field(50015; "Shading Style"; Option)
        {
            Caption = 'Shading Style';
            DataClassification = ToBeClassified;
            OptionCaption = 'None,Horizontal,Horizontal Middle,Vertical,Vertical Middle,Diagonal Up,Diagonal Up Middle,Diagonal Down,Diagonal Down Middle,From Left Top Corner,From Right Top Corner,From Left Bottom Corner,From Right Bottom Corner,From Center';
            OptionMembers = None,Horizontal,"Horizontal Middle",Vertical,"Vertical Middle","Diagonal Up","Diagonal Up Middle","Diagonal Down","Diagonal Down Middle","From Left Top Corner","From Right Top Corner","From Left Bottom Corner","From Right Bottom Corner","From Center";
        }
        field(50020; "Border Style"; Enum "Excel Buffer Border Style")
        {
            Caption = 'Border Style';
            DataClassification = ToBeClassified;
        }
        field(50021; "Border Color"; Text[30])
        {
            Caption = 'Border Color';
            DataClassification = ToBeClassified;
        }
        field(50022; "Left Border Style"; Enum "Excel Buffer Border Style")
        {
            Caption = 'Left Border Style';
            DataClassification = ToBeClassified;
        }
        field(50023; "Left Border Color"; Text[30])
        {
            Caption = 'Left Border Color';
            DataClassification = ToBeClassified;
        }
        field(50024; "Right Border Style"; Enum "Excel Buffer Border Style")
        {
            Caption = 'Right Border Style';
            DataClassification = ToBeClassified;
        }
        field(50025; "Right Border Color"; Text[30])
        {
            Caption = 'Right Border Color';
            DataClassification = ToBeClassified;
        }
        field(50026; "Top Border Style"; Enum "Excel Buffer Border Style")
        {
            Caption = 'Top Border Style';
            DataClassification = ToBeClassified;
        }
        field(50027; "Top Border Color"; Text[30])
        {
            Caption = 'Top Border Color';
            DataClassification = ToBeClassified;
        }
        field(50028; "Bottom Border Style"; Enum "Excel Buffer Border Style")
        {
            Caption = 'Bottom Border Style';
            DataClassification = ToBeClassified;
        }
        field(50029; "Bottom Border Color"; Text[30])
        {
            Caption = 'Bottom Border Color';
            DataClassification = ToBeClassified;
        }
        field(50030; "Diagonal Border Style"; Enum "Excel Buffer Border Style")
        {
            Caption = 'Diagonal Border Up Style';
            DataClassification = ToBeClassified;
        }
        field(50031; "Diagonal Border Color"; Text[30])
        {
            Caption = 'Diagonal Border Up Color';
            DataClassification = ToBeClassified;
        }
        field(50032; "Diagonal Border Type"; Option)
        {
            Caption = 'Diagonal Border Type';
            DataClassification = ToBeClassified;
            OptionCaption = 'Up,Down,Up and Down';
            OptionMembers = Up,Down,"Up and Down";
        }

        field(50040; "Horizontal Alignment"; Option)
        {
            Caption = 'Horizontal Alignment';
            DataClassification = ToBeClassified;
            OptionCaption = 'Automatic,Left,Center,Right,Fill,Justify,CenterAcrossSelection,Distributed,JustifyDistributed';
            OptionMembers = Automatic,Left,Center,Right,Fill,Justify,CenterAcrossSelection,Distributed,JustifyDistributed;
        }
        field(50041; "Vertical Alignment"; Option)
        {
            Caption = 'Vertical Alignment';
            DataClassification = ToBeClassified;
            OptionCaption = 'Automatic,Top,Center,Bottom,Justify,Distributed,JustifyDistributed';
            OptionMembers = Automatic,Top,Center,Bottom,Justify,Distributed,JustifyDistributed;
        }
        field(50042; Indent; Integer)
        {
            Caption = 'Indent';
            DataClassification = ToBeClassified;
        }
        field(50043; "Reading Order"; Option)
        {
            Caption = 'Reading Order';
            DataClassification = ToBeClassified;
            OptionCaption = 'Context,Left-to-Right,Right-to-Left';
            OptionMembers = Context,"Left-to-Right","Right-to-Left";
        }
        field(50044; "Relative Indent"; Integer)
        {
            Caption = 'Relative Indent';
            DataClassification = ToBeClassified;
        }
        field(50045; "Shrink To Fit"; Boolean)
        {
            Caption = 'Shrink To Fit';
            DataClassification = ToBeClassified;
        }
        field(50046; "Text Rotation"; Integer)
        {
            Caption = 'Text Rotation';
            DataClassification = ToBeClassified;
            MinValue = 0;
            MaxValue = 180;
        }
        field(50047; "Wrap Text"; Boolean)
        {
            Caption = 'Wrap Text';
            DataClassification = ToBeClassified;
        }
        field(50048; "Justify Last Line"; Boolean)
        {
            Caption = 'Justify Last Line';
            DataClassification = ToBeClassified;
        }
        field(50100; "Column Width"; Integer)
        {
            Caption = 'Column Width';
            DataClassification = ToBeClassified;
        }

        field(50101; "Column Hidden"; Boolean)
        {
            Caption = 'Column Hidden';
            DataClassification = ToBeClassified;
        }
        field(50102; "Column Outline Level"; Integer)
        {
            Caption = 'Column Outline Level';
            DataClassification = ToBeClassified;
        }
        field(50110; "Row Height"; Integer)
        {
            Caption = 'Row Height';
            DataClassification = ToBeClassified;
        }

        field(50111; "Row Hidden"; Boolean)
        {
            Caption = 'Row Hidden';
            DataClassification = ToBeClassified;
        }
        field(50112; "Row Outline Level"; Integer)
        {
            Caption = 'Row Outline Level';
            DataClassification = ToBeClassified;
        }
        field(50113; "Row Collapsed"; Boolean)
        {
            Caption = 'Row Collapsed';
            DataClassification = ToBeClassified;
        }
    }

    keys
    {
        key(Key1; "Row No.", "Column No.")
        {
            Clustered = true;
        }
    }

    fieldgroups
    {
    }

    var
        TempDefaultsExcelBuf: Record "Excel Buffer Extended" temporary;
        TempColumnsExcelBuf: Record "Excel Buffer Extended" temporary;
        TempRowsExcelBuf: Record "Excel Buffer Extended" temporary;
        FileManagement: Codeunit "File Management";
        ExcelWorkbook: Codeunit "Excel Workbook";
        ExcelWorksheet: Codeunit "Excel Worksheet";
        RangeStartXlRow: Text[30];
        RangeStartXlCol: Text[30];
        RangeEndXlRow: Text[30];
        RangeEndXlCol: Text[30];
        FileNameServer: Text;
        FriendlyName: Text;
        CurrentRow: Integer;
        CurrentCol: Integer;
        ErrorMessage: Text;
        ColumnProperty: Option Width,OutlineLevel,Hidden;
        RowProperty: Option Height,OutlineLevel,Hidden,Collapsed;
        Text001: Label 'You must enter a file name.';
        Text002: Label 'You must enter an Excel worksheet name.', Comment = '{Locked="Excel"}';
        Text003: Label 'The file %1 does not exist.';
        Text004: Label 'The Excel worksheet %1 does not exist.', Comment = '{Locked="Excel"}';
        Text005: Label 'Creating Excel worksheet...\\', Comment = '{Locked="Excel"}';
        PageTxt: Label 'Page';
        Text007: Label 'Reading Excel worksheet...\\', Comment = '{Locked="Excel"}';
        Text013: Label '&B';
        Text014: Label '&D';
        Text015: Label '&P';
        Text016: Label 'A1';
        Text017: Label 'SUMIF';
        Text018: Label '#N/A';
        Text019: Label 'GLAcc', Comment = 'Used to define an Excel range name. You must refer to Excel rules to change this term.', Locked = true;
        Text020: Label 'Period', Comment = 'Used to define an Excel range name. You must refer to Excel rules to change this term.', Locked = true;
        Text021: Label 'Budget';
        Text022: Label 'CostAcc', Locked = true, Comment = 'Used to define an Excel range name. You must refer to Excel rules to change this term.';
        Text034: Label 'Excel Files (*.xls*)|*.xls*|All Files (*.*)|*.*', Comment = '{Split=r''\|\*\..{1,4}\|?''}{Locked="Excel"}';
        Text035: Label 'The operation was canceled.';
        Text037: Label 'Could not create the Excel workbook.', Comment = '{Locked="Excel"}';
        Text038: Label 'Global variable %1 is not included for test.';
        Text039: Label 'Cell type has not been set.';
        SavingDocumentMsg: Label 'Saving the following document: %1.';
        ExcelFileExtensionTok: Label '.xlsx', Locked = true;
        CellNotFoundErr: Label 'Cell %1 not found.', Comment = '%1 - cell name';
        InvalidValueErr: Label 'The parameter ''%1'' has invalid value %2.';
        InvalidRangeErr: Label '%1 is invalid range.';



    procedure CreateBook(SheetName: Text)
    begin
        ExcelWorkbook.CreateBook(SheetName, ExcelWorksheet);
    end;

    procedure CloseBook()
    begin
        ExcelWorkbook.CloseBook();
    end;

    procedure CloseAndDownloadBook(Name: text)
    begin
        CloseBook();
        SetFriendlyFilename(Name);
        DownloadExcelFile();
    end;

    procedure AddNewSheet(NewSheetName: Text)
    begin
        ExcelWorkbook.AddNewSheet(NewSheetName, ExcelWorksheet);
        ClearNewRow();
    end;

    procedure WriteSheet(ReportHeader: Text; CompanyName2: Text; UserID2: Text)
    var
        TypeHelper: Codeunit "Type Helper";
        PageHeader: Text;
    begin
        /*
        XlWrkShtWriter.AddPageSetup(OrientationValues.Landscape, 9); // 9 - default value for Paper Size - A4
        if ReportHeader <> '' then
            XlWrkShtWriter.AddHeader(
              true,
              StrSubstNo('%1%2%1%3%4', GetExcelReference(1), ReportHeader, TypeHelper.LFSeparator(), CompanyName2));

        XlWrkShtWriter.AddHeader(
          false,
          StrSubstNo('%1%3%4%3%5 %2', GetExcelReference(2), GetExcelReference(3), TypeHelper.LFSeparator(), UserID2, PageTxt));
        
        OpenXMLManagement.AddAndInitializeCommentsPart(XlWrkShtWriter, VmlDrawingPart);

        StringBld := StringBld.StringBuilder();
        StringBld.Append(VmlDrawingXmlTxt);
        */

        if ReportHeader <> '' then
            PageHeader :=
              StrSubstNo('&C%1%2%1%3%4', GetExcelReference(1), ReportHeader, TypeHelper.LFSeparator(), CompanyName2);
        if UserID2 <> '' then
            PageHeader +=
                StrSubstNo('&R%1%3%4%3%5 %2', GetExcelReference(2), GetExcelReference(3), TypeHelper.LFSeparator(), UserID2, PageTxt);
        if PageHeader <> '' then
            ExcelWorksheet.AddPageHeaderFooter("Excel Page HeaderFooter Type"::oddHeader, PageHeader);

        //&amp;LEven page left &amp;Ceven page center&amp;REven page right

        ExcelWorksheet.WriteSheetData(Rec, TempRowsExcelBuf);
        ExcelWorksheet.WriteColumnsProperties(TempColumnsExcelBuf);
        ExcelWorkbook.SaveWorksheet(ExcelWorksheet);

        TempColumnsExcelBuf.Reset;
        TempColumnsExcelBuf.DeleteAll;

        TempRowsExcelBuf.Reset;
        TempRowsExcelBuf.DeleteAll;

        /*
        StringBld.Append(EndXmlTokenTxt);

        IsHandled := false;
        OnWriteSheetOnBeforeUseXmlTextWriter(Rec, IsHandled);
        if not IsHandled then begin
            XmlTextWriter := XmlTextWriter.XmlTextWriter(VmlDrawingPart.GetStream(FileMode.Create), Encoding.UTF8);
            XmlTextWriter.WriteRaw(StringBld.ToString());
            XmlTextWriter.Flush();
            XmlTextWriter.Close();
        end;        
        */
    end;

    procedure AddAutoFilter(RangeName: Text)
    begin
        ExcelWorksheet.AddAutoFilter(RangeName);
    end;

    procedure SetTabColor(NewTabColor: Text)
    begin
        ExcelWorksheet.SetTabColor(NewTabColor);
    end;

    procedure MergeCells(RangeName: Text)
    begin
        ExcelWorksheet.MergeCells(RangeName);
    end;

    procedure SetColumnWidth(ColNo: Integer; NewColWidth: Integer)
    begin
        SetColumnsProperty(ColNo, ColNo, ColumnProperty::Width, NewColWidth);
    end;

    procedure SetColumnsWidth(RangeName: Text; NewColWidth: Integer)
    begin
        SetColumnsProperty(RangeName, ColumnProperty::Width, NewColWidth);
    end;

    procedure SetColumnOutlineLevel(ColNo: Integer; NewOutlineLevel: Integer)
    begin
        SetColumnsProperty(ColNo, ColNo, ColumnProperty::OutlineLevel, NewOutlineLevel);
    end;

    procedure SetColumnsOutlineLevel(RangeName: Text; NewOutlineLevel: Integer)
    begin
        SetColumnsProperty(RangeName, ColumnProperty::OutlineLevel, NewOutlineLevel);
    end;

    procedure SetColumnHidden(ColNo: Integer; NewValue: Boolean)
    begin
        SetColumnsProperty(ColNo, ColNo, ColumnProperty::Hidden, NewValue);
    end;

    procedure SetColumnsHidden(RangeName: Text; NewValue: Boolean)
    begin
        SetColumnsProperty(RangeName, ColumnProperty::Hidden, NewValue);
    end;

    local procedure GetColumnRecord(ColumnNo: Integer);
    begin
        if not TempColumnsExcelBuf.Get(0, ColumnNo) then begin
            TempColumnsExcelBuf.Init;
            TempColumnsExcelBuf."Row No." := 0;
            TempColumnsExcelBuf.Validate("Column No.", ColumnNo);
            TempColumnsExcelBuf.Insert;
        end;
    end;

    local procedure SetColumnsProperty(RangeName: Text; PropertyType: Option Width,OutlineLevel,Hidden; Value: Variant);
    var
        ExcelBufferHelper: Codeunit "Excel Buffer Helper";
        FromRowNo: Integer;
        ToRowNo: Integer;
        FromColumnNo: Integer;
        ToColumnNo: Integer;
    begin
        if not ExcelBufferHelper.GetIntegerRange(RangeName, FromRowNo, FromColumnNo, ToRowNo, ToColumnNo) then
            Error(InvalidValueErr, 'RangeName', RangeName);

        if (FromColumnNo = 0) and (ToColumnNo = 0) then begin
            FromColumnNo := FromRowNo;
            ToColumnNo := ToRowNo;
        end;
        SetColumnsProperty(FromColumnNo, ToColumnNo, PropertyType, Value);
    end;

    local procedure SetColumnsProperty(FromColumnNo: Integer; ToColumnNo: Integer; PropertyType: Option Width,OutlineLevel,Hidden; Value: Variant);
    var
        I: Integer;
    begin
        if FromColumnNo < 1 then
            exit;

        for I := FromColumnNo to ToColumnNo do begin
            GetColumnRecord(I);
            case PropertyType of
                PropertyType::Width:
                    TempColumnsExcelBuf."Column Width" := Value;
                PropertyType::OutlineLevel:
                    TempColumnsExcelBuf."Column Outline Level" := Value;
                PropertyType::Hidden:
                    TempColumnsExcelBuf."Column Hidden" := Value;
            end;
            TempColumnsExcelBuf.Modify;
        end;
    end;

    procedure SetColumnsSummaryToLeft(Value: Boolean);
    begin
        ExcelWorksheet.SetColumnsSummaryToLeft(Value);
    end;

    local procedure GetRowRecord(RowNo: Integer);
    begin
        if not TempRowsExcelBuf.Get(RowNo, 0) then begin
            TempRowsExcelBuf.Init;
            TempRowsExcelBuf."Row No." := RowNo;
            TempRowsExcelBuf."Column No." := 0;
            TempRowsExcelBuf.Insert;
        end;
    end;

    procedure SetRowHeight(RowNo: Integer; NewRowHeight: Integer)
    begin
        SetRowsProperty(RowNo, RowNo, RowProperty::Height, NewRowHeight);
    end;

    procedure SetRowsHeight(FromRowNo: Integer; ToRowNo: Integer; NewRowHeight: Integer)
    begin
        SetRowsProperty(FromRowNo, ToRowNo, RowProperty::Height, NewRowHeight);
    end;

    procedure SetRowOutlineLevel(RowNo: Integer; NewOutlineLevel: Integer)
    begin
        SetRowsProperty(RowNo, RowNo, RowProperty::OutlineLevel, NewOutlineLevel);
    end;

    procedure SetRowsOutlineLevel(FromRowNo: Integer; ToRowNo: Integer; NewOutlineLevel: Integer)
    begin
        SetRowsProperty(FromRowNo, ToRowNo, RowProperty::OutlineLevel, NewOutlineLevel);
    end;

    procedure SetRowHidden(RowNo: Integer; NewValue: Boolean)
    begin
        SetRowsProperty(RowNo, RowNo, RowProperty::Hidden, NewValue);
    end;

    procedure SetRowsHidden(FromRowNo: Integer; ToRowNo: Integer; NewValue: Boolean)
    begin
        SetRowsProperty(FromRowNo, ToRowNo, RowProperty::Hidden, NewValue);
    end;

    procedure SetRowCollapsed(RowNo: Integer; NewValue: Boolean)
    begin
        SetRowsProperty(RowNo, RowNo, RowProperty::Collapsed, NewValue);
    end;

    procedure SetRowsCollapsed(FromRowNo: Integer; ToRowNo: Integer; NewValue: Boolean)
    begin
        SetRowsProperty(FromRowNo, ToRowNo, RowProperty::Collapsed, NewValue);
    end;

    local procedure SetRowsProperty(FromRowNo: Integer; ToRowNo: Integer; PropertyType: Option Height,OutlineLevel,Hidden,Collapsed; Value: Variant);
    var
        I: Integer;
    begin

        if FromRowNo < 1 then
            exit;

        for I := FromRowNo to ToRowNo do begin
            GetRowRecord(I);
            case PropertyType of
                PropertyType::Height:
                    TempRowsExcelBuf."Row Height" := Value;
                PropertyType::OutlineLevel:
                    TempRowsExcelBuf."Row Outline Level" := Value;
                PropertyType::Hidden:
                    TempRowsExcelBuf."Row Hidden" := Value;
                PropertyType::Collapsed:
                    TempRowsExcelBuf."Row Collapsed" := Value;
            end;
            TempRowsExcelBuf.Modify;
        end;
    end;

    procedure SetRowsSummaryAbove(Value: Boolean);
    begin
        ExcelWorksheet.SetRowsSummaryAbove(Value);
    end;

    [Scope('OnPrem')]
    procedure CreateRangeName(RangeName: Text[30]; FromColumnNo: Integer; FromRowNo: Integer)
    var
        TempExcelBuf: Record "Excel Buffer" temporary;
        ToxlRowID: Text[10];
    begin
        SetCurrentKey("Row No.", "Column No.");
        if Find('+') then
            ToxlRowID := xlRowID;
        TempExcelBuf.Validate("Row No.", FromRowNo);
        TempExcelBuf.Validate("Column No.", FromColumnNo);
        /*
        XlWrkShtWriter.AddRange(
          RangeName,
          GetExcelReference(4) + TempExcelBuf.xlColID + GetExcelReference(4) + TempExcelBuf.xlRowID +
          ':' +
          GetExcelReference(4) + TempExcelBuf.xlColID + GetExcelReference(4) + ToxlRowID);
        */
    end;


    procedure GetExcelReference(Which: Integer): Text[250]
    begin
        case Which of
            1:
                exit(Text013);
            // DO NOT TRANSLATE: &B is the Excel code to turn bold printing on or off for customized Header/Footer.
            2:
                exit(Text014);
            // DO NOT TRANSLATE: &D is the Excel code to print the current date in customized Header/Footer.
            3:
                exit(Text015);
            // DO NOT TRANSLATE: &P is the Excel code to print the page number in customized Header/Footer.
            4:
                exit('$');
            // DO NOT TRANSLATE: $ is the Excel code for absolute reference to cells.
            5:
                exit(Text016);
            // DO NOT TRANSLATE: A1 is the Excel reference of the first cell.
            6:
                exit(Text017);
            // DO NOT TRANSLATE: SUMIF is the name of the Excel function used to summarize values according to some conditions.
            7:
                exit(Text018);
            // DO NOT TRANSLATE: The #N/A Excel error value occurs when a value is not available to a function or formula.
            8:
                exit(Text019);
            // DO NOT TRANSLATE: GLAcc is used to define an Excel range name. You must refer to Excel rules to change this term.
            9:
                exit(Text020);
            // DO NOT TRANSLATE: Period is used to define an Excel range name. You must refer to Excel rules to change this term.
            10:
                exit(Text021);
            // DO NOT TRANSLATE: Budget is used to define an Excel worksheet name. You must refer to Excel rules to change this term.
            11:
                exit(Text022);
        // DO NOT TRANSLATE: CostAcc is used to define an Excel range name. You must refer to Excel rules to change this term.
        end;
    end;

    procedure AddToFormula(Text: Text[30]): Boolean
    var
        Overflow: Boolean;
        LongFormula: Text[1000];
    begin
        LongFormula := GetFormula();
        if LongFormula = '' then
            LongFormula := '=';
        if LongFormula <> '=' then
            if StrLen(LongFormula) + 1 > MaxStrLen(LongFormula) then
                Overflow := true
            else
                LongFormula := LongFormula + '+';
        if StrLen(LongFormula) + StrLen(Text) > MaxStrLen(LongFormula) then
            Overflow := true
        else
            SetFormula(LongFormula + Text);
        exit(Overflow);
    end;

    procedure GetFormula(): Text[2000]
    begin
        exit(Formula + Formula2 + Formula3 + Formula4);
    end;

    procedure SetFormula(LongFormula: Text[2000])
    begin
        ClearFormula();
        if LongFormula = '' then
            exit;

        Formula := CopyStr(LongFormula, 1, MaxStrLen(Formula));
        if StrLen(LongFormula) > MaxStrLen(Formula) then
            Formula2 := CopyStr(LongFormula, MaxStrLen(Formula) + 1, MaxStrLen(Formula2));
        if StrLen(LongFormula) > MaxStrLen(Formula) + MaxStrLen(Formula2) then
            Formula3 := CopyStr(LongFormula, MaxStrLen(Formula) + MaxStrLen(Formula2) + 1, MaxStrLen(Formula3));
        if StrLen(LongFormula) > MaxStrLen(Formula) + MaxStrLen(Formula2) + MaxStrLen(Formula3) then
            Formula4 := CopyStr(LongFormula, MaxStrLen(Formula) + MaxStrLen(Formula2) + MaxStrLen(Formula3) + 1, MaxStrLen(Formula4));
    end;

    procedure ClearFormula()
    begin
        Formula := '';
        Formula2 := '';
        Formula3 := '';
        Formula4 := '';
    end;

    procedure NewRow()
    begin
        SetCurrent(CurrentRow + 1, 0);
    end;

    procedure AddColumn(Value: Variant; IsFormula: Boolean; CommentText: Text; IsBold: Boolean; IsItalics: Boolean; IsUnderline: Boolean; NumFormat: Text[30]; CellType: Option)
    begin
        if CurrentRow < 1 then
            NewRow();

        CurrentCol := CurrentCol + 1;
        Init();
        TransferFields(TempDefaultsExcelBuf, false);
        Validate("Row No.", CurrentRow);
        Validate("Column No.", CurrentCol);
        if IsFormula then
            SetFormula(Format(Value))
        else
            SetCellValue(Format(Value));
        Comment := CopyStr(CommentText, 1, MaxStrLen(Comment));
        Bold := IsBold;
        Italic := IsItalics;
        Underline := IsUnderline;
        NumberFormat := NumFormat;
        "Cell Type" := CellType;
        Insert();
    end;

    procedure AddColumn(Value: Variant; IsFormula: Boolean; CellType: Option; Properties: Text);
    begin
        if CurrentRow < 1 then
            NewRow;
        CurrentCol += 1;

        Init();
        TransferFields(TempDefaultsExcelBuf, false);
        Validate("Row No.", CurrentRow);
        Validate("Column No.", CurrentCol);
        if IsFormula then
            SetFormula(Format(Value))
        else
            SetCellValue(Format(Value));
        "Cell Type" := CellType;
        SetProperties(Properties);
        Insert();

        if "Column Width" > 0 then
            SetColumnWidth(CurrentCol, "Column Width");
        if "Column Outline Level" > 0 then
            SetColumnOutlineLevel(CurrentCol, "Column Outline Level");
        if "Column Hidden" then
            SetColumnHidden(CurrentCol, "Column Hidden");

        if "Row Height" > 0 then
            SetRowHeight(CurrentRow, "Row Height");
        if "Row Hidden" then
            SetRowHidden(CurrentRow, true);
        if "Row Outline Level" > 0 then
            SetRowOutlineLevel(CurrentRow, "Row Outline Level");
        if "Row Collapsed" then
            SetRowCollapsed(CurrentRow, "Row Collapsed");
    end;

    procedure AddColumnWithLink(Value: Variant; Link: Text; CellType: Option; Properties: Text);
    var
        TempDefaultsExcelBuf2: Record "Excel Buffer Extended";
    begin
        TempDefaultsExcelBuf2 := TempDefaultsExcelBuf;
        TempDefaultsExcelBuf.Underline := true;
        TempDefaultsExcelBuf."Font Color" := 'FF0000FF';

        AddColumn(
          StrSubstNo('=HYPERLINK("%1","%2")', Link, Value),
          true,
          CellType,
          Properties);

        TempDefaultsExcelBuf := TempDefaultsExcelBuf2;
    end;

    procedure AddColumnAndMerge(Value: Variant; IsFormula: Boolean; CellType: Option; Properties: Text; MergeCols: Integer);
    var
        ExcelBufferHelper: Codeunit "Excel Buffer Helper";
    begin
        AddColumn(Value, IsFormula, CellType, Properties);

        if MergeCols = 0 then
            exit;

        if MergeCols > 0 then begin
            CurrentCol += MergeCols;

            MergeCells(
              StrSubstNo('%1%2:%3%4', xlColID, xlRowID, ExcelBufferHelper.GetColumnCode(CurrentCol), xlRowID))
        end else begin
            if CurrentCol <= -MergeCols then
                MergeCols := -(CurrentCol - 1);

            MergeCells(
              StrSubstNo('%1%2:%3%4', ExcelBufferHelper.GetColumnCode(CurrentCol + MergeCols), xlRowID, xlColID, xlRowID));
        end;
    end;

    procedure EnterCell(RowNo: Integer; ColumnNo: Integer; Value: Variant; IsBold: Boolean; IsItalics: Boolean; IsUnderline: Boolean)
    begin
        Init();
        Validate("Row No.", RowNo);
        Validate("Column No.", ColumnNo);

        case true of
            Value.IsDecimal or Value.IsInteger or Value.IsBigInteger:
                Validate("Cell Type", "Cell Type"::Number);
            Value.IsDate or Value.IsDateTime:
                Validate("Cell Type", "Cell Type"::Date);
            Value.IsTime:
                Validate("Cell Type", "Cell Type"::Time);
            else
                Validate("Cell Type", "Cell Type"::Text);
        end;

        "Cell Value as Text" := CopyStr(Format(Value), 1, MaxStrLen("Cell Value as Text"));
        Bold := IsBold;
        Italic := IsItalics;
        Underline := IsUnderline;
        Insert(true);
    end;

    procedure StartRange()
    var
        DummyExcelBuf: Record "Excel Buffer";
    begin
        Error('This function is not implemented');
        DummyExcelBuf.Validate("Row No.", CurrentRow);
        DummyExcelBuf.Validate("Column No.", CurrentCol);

        RangeStartXlRow := DummyExcelBuf.xlRowID;
        RangeStartXlCol := DummyExcelBuf.xlColID;
    end;

    procedure EndRange()
    var
        DummyExcelBuf: Record "Excel Buffer";
    begin
        Error('This function is not implemented');
        DummyExcelBuf.Validate("Row No.", CurrentRow);
        DummyExcelBuf.Validate("Column No.", CurrentCol);

        RangeEndXlRow := DummyExcelBuf.xlRowID;
        RangeEndXlCol := DummyExcelBuf.xlColID;
    end;

    procedure CreateRange(RangeName: Text[250])
    begin
        Error('This function is not implemented');
        /*
        XlWrkShtWriter.AddRange(
          RangeName,
          GetExcelReference(4) + RangeStartXlCol + GetExcelReference(4) + RangeStartXlRow +
          ':' +
          GetExcelReference(4) + RangeEndXlCol + GetExcelReference(4) + RangeEndXlRow);
        */
    end;

    procedure BorderAround(BorderStyle: Enum "Excel Buffer Border Style");
    var
    begin
        BorderAround(Rec."Row No.", Rec."Column No.", Rec."Row No.", Rec."Column No.", BorderStyle);
    end;

    procedure BorderAround(RangeName: Text; BorderStyle: Enum "Excel Buffer Border Style");
    var
        ExcelBufferHelper: Codeunit "Excel Buffer Helper";
        FromRowNo: Integer;
        ToRowNo: Integer;
        FromColumnNo: Integer;
        ToColumnNo: Integer;
    begin
        if not ExcelBufferHelper.GetIntegerRange(RangeName, FromRowNo, FromColumnNo, ToRowNo, ToColumnNo) then
            Error(InvalidValueErr, 'RangeName', RangeName);
        BorderAround(FromRowNo, FromColumnNo, ToRowNo, ToColumnNo, BorderStyle);
    end;

    procedure BorderAround(FromRowNo: Integer; FromColumnNo: Integer; ToRowNo: Integer; ToColumnNo: Integer; BorderStyle: Enum "Excel Buffer Border Style");
    var
        I: Integer;
        J: Integer;
        RecordExists: Boolean;
    begin
        if (FromRowNo < 1) or (FromColumnNo < 1) then
            exit;
        if (FromRowNo > ToRowNo) or (FromColumnNo > ToColumnNo) then
            exit;

        for I := FromRowNo to ToRowNo do
            for J := FromColumnNo to ToColumnNo do begin
                RecordExists := Rec.Get(I, J);
                if not RecordExists then begin
                    Rec.Init;
                    Rec.Validate("Row No.", I);
                    Rec.Validate("Column No.", J);
                end;
                if I = FromRowNo then
                    "Top Border Style" := BorderStyle;
                if I = ToRowNo then
                    "Bottom Border Style" := BorderStyle;
                if J = FromColumnNo then
                    "Left Border Style" := BorderStyle;
                if J = ToColumnNo then
                    "Right Border Style" := BorderStyle;

                if RecordExists then
                    Rec.Modify
                else
                    Rec.Insert;
                if not (I in [FromRowNo, ToRowNo]) and (J = FromColumnNo) and (J < ToColumnNo) then
                    J := ToColumnNo - 1;
            end;
    end;

    procedure ClearNewRow()
    begin
        SetCurrent(0, 0);
    end;

    procedure UTgetGlobalValue(globalVariable: Text[30]; var value: Variant)
    begin
        case globalVariable of
            'CurrentRow':
                value := CurrentRow;
            'CurrentCol':
                value := CurrentCol;
            'RangeStartXlRow':
                value := RangeStartXlRow;
            'RangeStartXlCol':
                value := RangeStartXlCol;
            'RangeEndXlRow':
                value := RangeEndXlRow;
            'RangeEndXlCol':
                value := RangeEndXlCol;
            //'XlWrkSht':
            //    value := XlWrkShtWriter;
            'ExcelFile':
                value := FileNameServer;
            else
                Error(Text038, globalVariable);
        end;
    end;

    procedure SetCurrent(NewCurrentRow: Integer; NewCurrentCol: Integer)
    begin
        CurrentRow := NewCurrentRow;
        CurrentCol := NewCurrentCol;
    end;

    procedure GetCurrentRow(): Integer
    begin
        exit(CurrentRow)
    end;

    procedure GetCurrentCol(): Integer
    begin
        exit(CurrentCol);
    end;

    procedure CreateValidationRule(Range: Code[20])
    begin
        Error('This function is not implemented');
        /*
        XlWrkShtWriter.AddRangeDataValidation(
          Range,
          GetExcelReference(4) + RangeStartXlCol + GetExcelReference(4) + RangeStartXlRow +
          ':' +
          GetExcelReference(4) + RangeEndXlCol + GetExcelReference(4) + RangeEndXlRow);
        */
    end;

    procedure QuitExcel()
    begin
        CloseBook();
    end;

    procedure DownloadExcelFile()
    begin
        ExcelWorkbook.DownloadBook(GetFriendlyFilename());
    end;

    local procedure GetFriendlyFilename(): Text
    begin
        if FriendlyName = '' then
            exit('Book1' + ExcelFileExtensionTok);

        exit(FileManagement.StripNotsupportChrInFileName(FriendlyName) + ExcelFileExtensionTok);
    end;

    procedure SetFriendlyFilename(Name: Text)
    begin
        FriendlyName := Name;
    end;

    procedure SaveToStream(var ResultStream: OutStream)
    begin
        ExcelWorkbook.SaveToStream(ResultStream);
    end;

    local procedure SetCellValue(NewValue: Text);
    var
        OutStream: OutStream;
    begin
        if StrLen(NewValue) <= MaxStrLen("Cell Value as Text") then begin
            "Cell Value as Text" := NewValue;
            exit;
        end;

        "Cell Value as Blob".CreateOutStream(OutStream, TextEncoding::Windows);
        OutStream.Write(NewValue);
    end;

    procedure SetDefaultProperties(DefaultValues: Record "Excel Buffer Extended");
    begin
        Clear(TempDefaultsExcelBuf);
        TempDefaultsExcelBuf := DefaultValues;
        InitTempDefaultsExcelBuf();
    end;

    procedure SetDefaultProperties(DefaultValues: Text);
    begin
        Clear(TempDefaultsExcelBuf);
        TempDefaultsExcelBuf.SetProperties(DefaultValues);
        InitTempDefaultsExcelBuf();
    end;

    local procedure InitTempDefaultsExcelBuf()
    begin
        TempDefaultsExcelBuf."Row No." := 0;
        TempDefaultsExcelBuf."Column No." := 0;
        TempDefaultsExcelBuf.xlRowID := '';
        TempDefaultsExcelBuf.xlColID := '';
        TempDefaultsExcelBuf."Cell Value as Text" := '';
        TempDefaultsExcelBuf.SetFormula('');
    end;

    procedure SetProperties(Properties: Text);
    var
        Field: Record Field;
        RecRef: RecordRef;
        FieldRef: FieldRef;
        Pos: Integer;
        Text: Text;
        FieldName: Text;
        FieldValue: Text;
    begin
        if Properties = '' then
            exit;

        RecRef.GetTable(Rec);

        while Properties <> '' do begin
            Pos := StrPos(Properties, ';');
            Text := '';
            if Pos > 0 then begin
                if Pos > 1 then
                    Text := CopyStr(Properties, 1, Pos - 1);
                Properties := CopyStr(Properties, Pos + 1);
            end else begin
                Text := Properties;
                Properties := '';
            end;

            while (Text <> '') and (Text[1] = ' ') do
                Text := CopyStr(Text, 2);

            while (Text <> '') and (Text[StrLen(Text)] = ' ') do
                Text := CopyStr(Text, 1, StrLen(Text) - 1);

            FieldName := '';
            FieldValue := '';
            Pos := StrPos(Text, ':');
            if Pos > 0 then begin
                FieldName := CopyStr(Text, 1, Pos - 1).Trim();
                FieldValue := CopyStr(Text, Pos + 1).Trim();
            end else
                FieldName := Text;

            Field.Reset();
            Field.SetRange(TableNo, RecRef.Number);
            Field.SetRange(FieldName, FieldName);
            if Field.FindFirst() then begin
                FieldRef := RecRef.Field(Field."No.");

                if (Field.Type = Field.Type::Boolean) and (FieldValue = '') then
                    FieldValue := Format(true, 0, 9);
                Evaluate(FieldRef, FieldValue, 9);
            end;
        end;

        RecRef.SetTable(Rec);
    end;

    procedure SetPaperSize(PaperSize: enum "Excel Paper Size"; Orientation: enum "Excel Page Orientation")
    begin
        ExcelWorksheet.SetPaperSize(PaperSize, Orientation);
    end;

    procedure SetPageScale(FitToPage: Enum "Excel Fit Page Scale")
    begin
        ExcelWorksheet.SetPageScale(FitToPage);
    end;

    procedure SetPageMargins(Left: Decimal; Top: Decimal; Right: Decimal; Bottom: Decimal; Header: Decimal; Footer: Decimal);
    begin
        ExcelWorksheet.SetPageMargins(Left, Top, Right, Bottom, Header, Footer);
    end;

    procedure SetPageMarginsNarrow();
    begin
        SetPageMargins(0.25, 0.75, 0.25, 0.75, 0.3, 0.3);
    end;

    procedure SetPageMarginsWide();
    begin
        SetPageMargins(1, 1, 1, 1, 0.5, 0.5);
    end;

    procedure AddFreezePane(FreezePaneRowNo: Integer; FreezePaneColNo: Integer);
    begin
        ExcelWorksheet.AddFreezePane(FreezePaneRowNo, FreezePaneColNo);
    end;

    procedure FreezeTopRow();
    begin
        AddFreezePane(2, 1);
    end;

    procedure FreezeFirstColumn();
    begin
        AddFreezePane(1, 2);
    end;

    procedure SetPageHeaderFooterSettings(DifferentOddAndEvenPages: Boolean; DifferentFirstPage: Boolean)
    begin
        ExcelWorksheet.SetPageHeaderFooterSettings(DifferentOddAndEvenPages, DifferentFirstPage);
    end;

    procedure SetPageHeaderFooterSettings(DifferentOddAndEvenPages: Boolean; DifferentFirstPage: Boolean; ScaleWithDoc: Boolean; AlignWithMargins: Boolean)
    begin
        ExcelWorksheet.SetPageHeaderFooterSettings(DifferentOddAndEvenPages, DifferentFirstPage, ScaleWithDoc, AlignWithMargins);
    end;

    procedure AddPageHeaderFooter(ValueType: Enum "Excel Page HeaderFooter Type"; Value: Text)
    begin
        ExcelWorksheet.AddPageHeaderFooter(ValueType, Value);
    end;

    procedure AddPageHeaderFooter(ValueType: Enum "Excel Page HeaderFooter Type"; LeftText: Text; CenterText: Text; RightText: Text)
    begin
        ExcelWorksheet.AddPageHeaderFooter(ValueType, StrSubstNo('&L%1&C%2&R%3', LeftText, CenterText, RightText));
    end;
}

