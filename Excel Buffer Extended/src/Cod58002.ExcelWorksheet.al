namespace ASD.Excel;
using System.IO;
using System.Utilities;
codeunit 58002 "Excel Worksheet"
{
    var
        ExcelBufferHelper: Codeunit "Excel Buffer Helper";
        WorkbookStyle: Codeunit "Excel Workbook Style";
        WorksheetDoc: XmlDocument;
        DrawinigDoc: XmlDocument;
        CommentsDoc: XmlDocument;
        RelationshipsDoc: XmlDocument;
        rId: Integer;
        DrawinigCreated: Boolean;
        CommentsCreated: Boolean;
        RelationshipCreated: Boolean;
        CouldNotCreateBookErr: Label 'Could not create the Excel workbook.';
        CreatingExcelWorksheetTxt: Label 'Creating Excel worksheet...\\', Comment = '{Locked="Excel"}';
        OperationCanceledErr: Label 'The operation was canceled.';
        ElementDoesNotExistErr: Label 'Element %1 does not exist in the Worksheet document.';


    procedure InitWorksheet()
    begin
        // create xl\worksheets\sheet1.xml
        XmlDocument.ReadFrom(
            '<x:worksheet ' +
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
                'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
                'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ' +
                'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
                'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac" ' +
                'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                '<x:sheetData />' +
                '<x:pageSetup paperSize="9" orientation="landscape" />' +
            //<x:headerFooter>
            //    <x:evenHeader>&amp;D Page &amp;P</x:evenHeader>
            //</x:headerFooter>
            '</x:worksheet>',
           WorksheetDoc);

        Clear(DrawinigDoc);
        DrawinigCreated := false;

        Clear(CommentsDoc);
        CommentsCreated := false;

        Clear(RelationshipsDoc);
        RelationshipCreated := false;
    end;

    procedure GetXmlDocument(): XmlDocument
    begin
        exit(WorksheetDoc);
    end;

    local procedure GetWorksheetChildElement(WkshtElementName: Text): XmlElement;
    var
        FoundXml: XmlElement;
    begin
        if ExcelBufferHelper.SelectSingleXmlElement(WorksheetDoc, '/x:worksheet/x:' + WkshtElementName, FoundXml) then
            exit(FoundXml);

        exit(CreateWorksheetChildElement(WkshtElementName));

    end;

    local procedure CreateWorksheetChildElement(NewElementName: Text) CreatedXml: XmlElement;
    var
        WkshtElements: List of [Text];
        WorksheetXml: XmlElement;
    begin
        WkshtElements.Add('sheetPr');
        WkshtElements.Add('dimension');
        WkshtElements.Add('sheetViews');
        WkshtElements.Add('sheetFormatPr');
        WkshtElements.Add('cols');
        WkshtElements.Add('sheetData');
        WkshtElements.Add('sheetCalcPr');
        WkshtElements.Add('sheetProtection');
        WkshtElements.Add('protectedRanges');
        WkshtElements.Add('scenarios');
        WkshtElements.Add('autoFilter');
        WkshtElements.Add('sortState');
        WkshtElements.Add('dataConsolidate');
        WkshtElements.Add('customSheetViews');
        WkshtElements.Add('mergeCells');
        WkshtElements.Add('phoneticPr');
        WkshtElements.Add('conditionalFormatting');
        WkshtElements.Add('dataValidations');
        WkshtElements.Add('hyperlinks');
        WkshtElements.Add('printOptions');
        WkshtElements.Add('pageMargins');
        WkshtElements.Add('pageSetup');
        WkshtElements.Add('headerFooter');
        WkshtElements.Add('rowBreaks');
        WkshtElements.Add('colBreaks');
        WkshtElements.Add('customProperties');
        WkshtElements.Add('cellWatches');
        WkshtElements.Add('ignoredErrors');
        WkshtElements.Add('smartTags');
        WkshtElements.Add('drawing');
        WkshtElements.Add('legacyDrawing');
        WkshtElements.Add('legacyDrawingHF');
        WkshtElements.Add('drawingHF');
        WkshtElements.Add('picture');
        WkshtElements.Add('oleObjects');
        WkshtElements.Add('controls');
        WkshtElements.Add('webPublishItems');
        WkshtElements.Add('tableParts');
        WkshtElements.Add('extLst');
        WorksheetDoc.GetRoot(WorksheetXml);
        Exit(ExcelBufferHelper.CreateChildElement(WkshtElements, WorksheetXml, NewElementName));
    end;

    procedure SetWorkbookStyle(var NewWorkbookStyle: Codeunit "Excel Workbook Style")
    begin
        WorkbookStyle := NewWorkbookStyle;
    end;

    procedure WriteSheetData(var ExcelBuffer: Record "Excel Buffer Extended"; var ExcelBufferRows: Record "Excel Buffer Extended");
    var
        ExcelBufferDialogMgt: Codeunit "Excel Buffer Dialog Management";
        SheetDataXml: XmlElement;
        RowXml: XmlElement;
        CellXml: XmlElement;
        MinRow: Integer;
        MaxRow: Integer;
        RowNo: Integer;
        RecNo: Integer;
        TotalRecNo: Integer;
        LastUpdate: DateTime;
        DialogOpened: Boolean;
        RowCreated: Boolean;
    begin
        if not ExcelBufferHelper.SelectSingleXmlElement(WorksheetDoc, '/x:worksheet/x:sheetData', SheetDataXml) then
            Error(CouldNotCreateBookErr);

        if ExcelBuffer.FindFirst() then
            MinRow := ExcelBuffer."Row No.";
        if ExcelBuffer.FindLast() then
            MaxRow := ExcelBuffer."Row No.";

        if ExcelBufferRows.FindFirst() then
            if MinRow > ExcelBufferRows."Row No." then
                MinRow := ExcelBufferRows."Row No.";
        if ExcelBufferRows.FindLast() then
            if MaxRow < ExcelBufferRows."Row No." then
                MaxRow := ExcelBufferRows."Row No.";

        TotalRecNo := MaxRow - MinRow;
        if TotalRecNo > 1 then begin
            ExcelBufferDialogMgt.Open(CreatingExcelWorksheetTxt);
            DialogOpened := true;
        end;
        LastUpdate := CurrentDateTime;

        for RowNo := MinRow to MaxRow do begin
            RowCreated := false;
            RowXml := XmlElement.Create('row', ExcelBufferHelper.MainNamespace);
            RowXml.SetAttribute('r', Format(RowNo, 0, 9));
            if ExcelBufferRows.Get(RowNo, 0) then begin
                if ExcelBufferRows."Row Height" > 0 then begin
                    RowXml.SetAttribute('customHeight', '1');
                    RowXml.SetAttribute('ht', Format(ExcelBufferRows."Row Height" * 0.75, 0, 9));
                end;

                if ExcelBufferRows."Row Hidden" then
                    RowXml.SetAttribute('hidden', '1');

                if ExcelBufferRows."Row Outline Level" > 0 then
                    RowXml.SetAttribute('outlineLevel', Format(ExcelBufferRows."Row Outline Level", 0, 9));

                if ExcelBufferRows."Row Collapsed" then
                    RowXml.SetAttribute('collapsed', '1');
                SheetDataXml.Add(RowXml);
                RowCreated := true;
            end;

            ExcelBuffer.SetRange("Row No.", RowNo);
            if ExcelBuffer.FindSet() then begin
                if not RowCreated then
                    SheetDataXml.Add(RowXml);
                repeat
                    WriteCellValue(ExcelBuffer, CellXml);
                    RowXml.Add(CellXml);
                    if ExcelBuffer.Comment <> '' then begin
                        SetCellComment(StrSubstNo('%1%2', ExcelBuffer.xlColID, ExcelBuffer."Row No."), ExcelBuffer.Comment);
                        AddCommentVmlShapeXml(ExcelBuffer."Column No.", ExcelBuffer."Row No.");
                    end;
                until ExcelBuffer.Next() = 0;
            end;

            ExcelBuffer.SetRange("Row No.");
        end;
    end;

    local procedure WriteCellValue(ExcelBuffer: Record "Excel Buffer Extended"; var CellXml: XmlElement)
    var
        CellTextValue: Text;
        RecInStream: InStream;
        DateValue: Date;
        DateTimeValue: DateTime;
        TimeValue: Time;
        DecValue: Decimal;
        StyleId: Integer;
    begin
        StyleId := WorkbookStyle.FindStyleId(ExcelBuffer);
        CellXml := XmlElement.Create('c', ExcelBufferHelper.MainNamespace);
        CellXml.SetAttribute('r', StrSubstNo('%1%2', ExcelBuffer.xlColID, ExcelBuffer."Row No.")); // Reference:	An A1 style reference to the location of this cell
        CellXml.SetAttribute('s', Format(StyleId, 0, 9));     // Style Index: The index of this cell's style. Style records are stored in the Styles Part.
        if ExcelBuffer.Formula <> '' then begin
            WriteFormula(CellXml, ExcelBuffer.GetFormula());
            exit;
        end;

        CellTextValue := ExcelBuffer."Cell Value as Text";

        if ExcelBuffer."Cell Value as Blob".HasValue() then begin
            ExcelBuffer.CalcFields("Cell Value as Blob");
            ExcelBuffer."Cell Value as Blob".CreateInStream(RecInStream, TextEncoding::Windows);
            RecInStream.ReadText(CellTextValue);
        end;

        case ExcelBuffer."Cell Type" of
            ExcelBuffer."Cell Type"::Number:
                if CellTextValue = '' then
                    WriteNumberValue(CellXml, '')
                else
                    if Evaluate(DecValue, CellTextValue) then
                        WriteNumberValue(CellXml, DecValue)
                    else
                        WriteTextValue(CellXml, CellTextValue);
            ExcelBuffer."Cell Type"::Text:
                WriteTextValue(CellXml, CellTextValue);
            ExcelBuffer."Cell Type"::Date:
                begin
                    if CellTextValue.Trim() = '' then begin
                        WriteNumberValue(CellXml, '');
                        exit;
                    end;

                    if Evaluate(DateTimeValue, CellTextValue) then begin
                        if DateTimeValue = 0DT then begin
                            WriteNumberValue(CellXml, '');
                            exit;
                        end;
                        DateValue := DT2Date(DateTimeValue);
                        TimeValue := DT2Time(DateTimeValue);

                        if DateValue < 19000101D then begin
                            WriteTextValue(CellXml, CellTextValue);
                            exit;
                        end;

                        if DateValue >= 19000301D then
                            DecValue := DateValue - 19000101D + 2
                        else
                            DecValue := DateValue - 19000101D + 1;
                        if TimeValue <> 0T then
                            DecValue += Round((TimeValue - 000000T) / (24 * 60 * 60 * 1000), 0.000000000000000001, '=');
                        WriteNumberValue(CellXml, DecValue);
                    end else
                        if Evaluate(DateValue, CellTextValue) then begin
                            if DateValue < 19000101D then begin
                                WriteTextValue(CellXml, CellTextValue);
                                exit;
                            end;
                            if DateValue >= 19000301D then
                                DecValue := DateValue - 19000101D + 2
                            else
                                DecValue := DateValue - 19000101D + 1;
                            WriteNumberValue(CellXml, DecValue)
                        end else
                            WriteTextValue(CellXml, CellTextValue);
                end;
            ExcelBuffer."Cell Type"::Time:
                begin
                    if CellTextValue.Trim() = '' then begin
                        WriteNumberValue(CellXml, '');
                        exit;
                    end;

                    if Evaluate(DateTimeValue, CellTextValue) then begin
                        TimeValue := DT2Time(DateTimeValue);
                        if TimeValue <> 0T then
                            DecValue := Round((TimeValue - 000000T) / (24 * 60 * 60 * 1000), 0.000000000000000001, '=');
                        WriteNumberValue(CellXml, DecValue);
                    end else
                        if Evaluate(TimeValue, CellTextValue) then begin
                            if TimeValue <> 0T then
                                DecValue := Round((TimeValue - 000000T) / (24 * 60 * 60 * 1000), 0.000000000000000001, '=');
                            WriteNumberValue(CellXml, DecValue);
                        end else
                            WriteTextValue(CellXml, CellTextValue);
                end;
        end;
    end;

    local procedure WriteTextValue(var CellXml: XmlElement; Text: Text)
    var
        InlineStrXml: XmlElement;
        CellValueXml: XmlElement;
    begin
        CellXml.SetAttribute('t', 'inlineStr');
        InlineStrXml := XmlElement.Create('is', ExcelBufferHelper.MainNamespace);
        CellValueXml := XmlElement.Create('t', ExcelBufferHelper.MainNamespace);
        CellValueXml.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve');
        CellValueXml.Add(XmlText.Create(Text));
        InlineStrXml.Add(CellValueXml);
        CellXml.Add(InlineStrXml);
    end;

    local procedure WriteNumberValue(var CellXml: XmlElement; Number: Decimal)
    begin
        WriteNumberValue(CellXml, Format(Number, 0, 9));
    end;

    local procedure WriteNumberValue(var CellXml: XmlElement; FormattedText: Text)
    var
        CellValueXml: XmlElement;
    begin
        CellXml.SetAttribute('t', 'n');
        CellValueXml := XmlElement.Create('v', ExcelBufferHelper.MainNamespace);
        CellValueXml.Add(XmlText.Create(FormattedText));
        CellXml.Add(CellValueXml);
    end;

    local procedure WriteFormula(var CellXml: XmlElement; Formula: Text)
    var
        CellValueXml: XmlElement;
    begin
        CellValueXml := XmlElement.Create('f', ExcelBufferHelper.MainNamespace);
        CellValueXml.Add(XmlText.Create(Formula));
        CellXml.Add(CellValueXml);
    end;

    procedure AddAutoFilter(RangeName: Text)
    var
        AutoFilterXml: XmlElement;
        SheetDataXml: XmlElement;
    begin
        AutoFilterXml := GetWorksheetChildElement('autoFilter');
        AutoFilterXml.SetAttribute('ref', RangeName);
    end;

    procedure MergeCells(RangeName: Text)
    var
        MergeCellsCollectionXml: XmlElement;
        MergeCellXml: XmlElement;
        PrevElement: XmlElement;
    begin
        MergeCellsCollectionXml := GetWorksheetChildElement('mergeCells');

        MergeCellXml := XmlElement.Create('mergeCell', ExcelBufferHelper.MainNamespace);
        MergeCellXml.SetAttribute('ref', RangeName);
        MergeCellsCollectionXml.Add(MergeCellXml);

        MergeCellsCollectionXml.SetAttribute('count', Format(MergeCellsCollectionXml.GetChildElements().Count, 0, 9));
    end;

    procedure WriteColumnsProperties(var TempColumnsExcelBuf: Record "Excel Buffer Extended" temporary);
    var
        TempColumnsExcelBuf2: Record "Excel Buffer Extended" temporary;
        ColsCollectionXml: XmlElement;
        ColumnXml: XmlElement;
        NextElement: XmlElement;
        MinColNo: Integer;
        MaxColNo: Integer;
        Pos: Integer;
        DoNext: Boolean;
    begin
        TempColumnsExcelBuf.Reset;
        if not TempColumnsExcelBuf.FindSet then
            exit;

        ColsCollectionXml := GetWorksheetChildElement('cols');

        repeat
            TempColumnsExcelBuf2 := TempColumnsExcelBuf;
            MinColNo := TempColumnsExcelBuf2."Column No.";
            MaxColNo := MinColNo;
            DoNext := TempColumnsExcelBuf.Next <> 0;
            while DoNext do begin
                if (TempColumnsExcelBuf."Column Width" = TempColumnsExcelBuf2."Column Width") and
                   (TempColumnsExcelBuf."Column Hidden" = TempColumnsExcelBuf2."Column Hidden") and
                   (TempColumnsExcelBuf."Column Outline Level" = TempColumnsExcelBuf2."Column Outline Level") and
                   (TempColumnsExcelBuf."Column No." = TempColumnsExcelBuf2."Column No." + 1)
                then begin
                    TempColumnsExcelBuf2 := TempColumnsExcelBuf;
                    MaxColNo := TempColumnsExcelBuf2."Column No.";
                    DoNext := TempColumnsExcelBuf.Next <> 0;
                end else
                    DoNext := false;
            end;
            TempColumnsExcelBuf := TempColumnsExcelBuf2;

            ColumnXml := XmlElement.Create('col', ExcelBufferHelper.MainNamespace);
            ColumnXml.SetAttribute('min', Format(MinColNo, 0, 9));
            ColumnXml.SetAttribute('max', Format(MaxColNo, 0, 9));

            if TempColumnsExcelBuf."Column Width" > 0 then begin
                ColumnXml.SetAttribute('width', Format(Round(TempColumnsExcelBuf."Column Width" / 7, 0.01, '='), 0, 9));
                ColumnXml.SetAttribute('customWidth', '1');
            end;

            if TempColumnsExcelBuf."Column Hidden" then
                ColumnXml.SetAttribute('hidden', '1');

            if TempColumnsExcelBuf."Column Outline Level" > 0 then begin
                ColumnXml.SetAttribute('outlineLevel', Format(TempColumnsExcelBuf."Column Outline Level", 0, 9));

                if TempColumnsExcelBuf."Column Width" = 0 then
                    ColumnXml.SetAttribute('width', Format(Round(64 / 7, 0.01, '='), 0, 9));
            end;

            if (TempColumnsExcelBuf."Column Width" > 0) or
               (TempColumnsExcelBuf."Column Hidden") or
               (TempColumnsExcelBuf."Column Outline Level" > 0)
            then
                ColsCollectionXml.Add(ColumnXml);

        until TempColumnsExcelBuf.Next = 0;
    end;

    procedure SetRowsSummaryAbove(Value: Boolean);
    var
        SheetPrXml: XmlElement;
        OutlinePrXml: XmlElement;
    begin
        SheetPrXml := GetWorksheetChildElement('sheetPr');

        if not ExcelBufferHelper.SelectSingleXmlElement(WorksheetDoc, '/x:worksheet/x:sheetPr/x:outlinePr', OutlinePrXml) then begin
            OutlinePrXml := XmlElement.Create('outlinePr', ExcelBufferHelper.MainNamespace);
            SheetPrXml.Add(OutlinePrXml);
        end;
        OutlinePrXml.SetAttribute('summaryBelow', Format(not Value, 0, 2));
    end;

    procedure SetColumnsSummaryToLeft(Value: Boolean);
    var
        SheetPrXml: XmlElement;
        OutlinePrXml: XmlElement;
    begin
        SheetPrXml := GetWorksheetChildElement('sheetPr');

        if not ExcelBufferHelper.SelectSingleXmlElement(WorksheetDoc, '/x:worksheet/x:sheetPr/x:outlinePr', OutlinePrXml) then begin
            OutlinePrXml := XmlElement.Create('outlinePr', ExcelBufferHelper.MainNamespace);
            SheetPrXml.Add(OutlinePrXml);
        end;
        OutlinePrXml.SetAttribute('summaryRight', Format(not Value, 0, 2));
    end;

    procedure SetTabColor(ColorName: Text);
    var
        SheetPrXml: XmlElement;
        TabColorXml: XmlElement;
    begin
        SheetPrXml := GetWorksheetChildElement('sheetPr');

        if not ExcelBufferHelper.SelectSingleXmlElement(WorksheetDoc, '/x:worksheet/x:sheetPr/x:tabColor', TabColorXml) then begin
            TabColorXml := XmlElement.Create('tabColor', ExcelBufferHelper.MainNamespace);
            SheetPrXml.Add(TabColorXml);
        end;
        TabColorXml.SetAttribute('rgb', ColorName);
    end;

    local procedure UpdateProgressDialog(var ExcelBufferDialogManagement: Codeunit "Excel Buffer Dialog Management"; var LastUpdate: DateTime; CurrentCount: Integer; TotalCount: Integer) Result: Boolean
    var
        CurrentTime: DateTime;
    begin

        // Refresh at 100%, and every second in between 0% to 100%
        // Duration is measured in miliseconds -> 1 sec = 1000 ms
        CurrentTime := CurrentDateTime;
        if (CurrentCount = TotalCount) or (CurrentTime - LastUpdate >= 1000) then begin
            LastUpdate := CurrentTime;
            if not ExcelBufferDialogManagement.SetProgress(Round(CurrentCount / TotalCount * 10000, 1)) then
                exit(false);
        end;

        exit(true)
    end;

    procedure SetCellComment(CellReference: Text; CommentValue: Text)
    var
        CommentListXml: XmlElement;
        CommentXml: XmlElement;
        TextElt: XmlElement;
        RunElt: XmlElement;
        RunPropElt: XmlElement;
        SpreadsheetTextElt: XmlElement;
    begin
        if not CommentsCreated then
            CreateComments();

        if not ExcelBufferHelper.SelectSingleXmlElement(CommentsDoc, '/x:comments/x:commentList', CommentListXml) then
            exit;

        CommentXml := XmlElement.Create('comment', ExcelBufferHelper.MainNamespace());
        CommentXml.SetAttribute('ref', CellReference);
        CommentXml.SetAttribute('authorId', '0');
        CommentXml.SetAttribute('shapeId', '0');
        CommentListXml.Add(CommentXml);

        TextElt := XmlElement.Create('text', ExcelBufferHelper.MainNamespace());
        CommentXml.Add(TextElt);

        RunElt := XmlElement.Create('r', ExcelBufferHelper.MainNamespace());
        TextElt.Add(RunElt);

        RunPropElt := XmlElement.Create('rPr', ExcelBufferHelper.MainNamespace());
        RunPropElt.Add(
            ExcelBufferHelper.CreateElementWithAttribute('sz', ExcelBufferHelper.MainNamespace(), 'val', '9'));
        RunPropElt.Add(
            ExcelBufferHelper.CreateElementWithAttribute('color', ExcelBufferHelper.MainNamespace(), 'indexed', '81'));
        RunPropElt.Add(
            ExcelBufferHelper.CreateElementWithAttribute('rFont', ExcelBufferHelper.MainNamespace(), 'val', 'Tahoma'));
        RunPropElt.Add(
            ExcelBufferHelper.CreateElementWithAttribute('charset', ExcelBufferHelper.MainNamespace(), 'val', '1'));
        RunElt.Add(RunPropElt);

        SpreadsheetTextElt := XmlElement.Create('t', ExcelBufferHelper.MainNamespace);
        SpreadsheetTextElt.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve');
        SpreadsheetTextElt.Add(XmlText.Create(CommentValue));
        RunElt.Add(SpreadsheetTextElt);

        /*
        <comment ref="D1" authorId="0" shapeId="0" xr:uid="{38A02300-D985-41F4-8912-A6E3CFCE8754}">
            <text>
                <r>
                    <rPr>
                        <sz val="9"/>
                        <color indexed="81"/>
                        <rFont val="Tahoma"/>
                        <charset val="1"/>
                    </rPr>
                    <t xml:space="preserve">It is a note</t>
                </r>
            </text>
        </comment>  
        */
    end;

    local procedure CreateDrawing()
    begin
        XmlDocument.ReadFrom(
            '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
                '<o:shapelayout v:ext="edit">' +
                    '<o:idmap v:ext="edit" data="1"/>' +
                '</o:shapelayout>' +
                '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"  path="m,l,21600r21600,l21600,xe">' +
                    '<v:stroke joinstyle="miter"/>' +
                    '<v:path gradientshapeok="t" o:connecttype="rect"/>' +
                '</v:shapetype>' +
            '</xml>',
            DrawinigDoc);
        DrawinigCreated := true;
    end;

    local procedure CreateComments()
    begin
        XmlDocument.ReadFrom(
            StrSubstNo(
                '<x:comments ' +
                    'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
                    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
                    'mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">' +
                    '<x:authors>' +
                        '<x:author>%1</x:author>' +
                    '</x:authors>' +
                    '<x:commentList />' +
                '</x:comments>', UserId),
            CommentsDoc);
        CommentsCreated := true;
    end;

    procedure AddCommentVmlShapeXml(ColId: Integer; RowId: Integer) CommentShape: Text
    var
        XmlDoc: XmlDocument;
        ShapeRootElement: XmlElement;
        DrawinigRootElement: XmlElement;
        XmlNode: XmlNode;
        Guid: Guid;
        Anchor: Text;
    begin
        if not DrawinigCreated then
            CreateDrawing();

        Guid := CreateGuid();
        Anchor := CreateCommentVmlAnchor(ColId, RowId);
        XmlDocument.ReadFrom(
            StrSubstNo(
                '<xml ' +
                    'xmlns:v="urn:schemas-microsoft-com:vml" ' +
                    'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
                    'xmlns:x="urn:schemas-microsoft-com:office:excel">' +
                    '<v:shape id="%1" type="#_x0000_t202" ' +
                        'style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:96pt;height:55.5pt;z-index:1;visibility:hidden" ' +
                        'fillcolor="#ffffe1" ' +
                        'o:insetmode="auto">' +
                        '<v:fill color2="#ffffe1" />' +
                        '<v:shadow color="black" obscured="t" />' +
                        '<v:path o:connecttype="none" />' +
                        '<v:textbox style="mso-direction-alt:auto">' +
                            '<div style="text-align:left" />' +
                        '</v:textbox>' +
                        '<x:ClientData ObjectType="Note">' +
                            '<x:MoveWithCells />' +
                            '<x:SizeWithCells />' +
                            '<x:Anchor>%2</x:Anchor>' +
                            '<x:AutoFill>False</x:AutoFill>' +
                            '<x:Row>%3</x:Row>' +
                            '<x:Column>%4</x:Column>' +
                        '</x:ClientData>' +
                    '</v:shape>' +
                '</xml>',
                Guid, Anchor, RowId - 1, ColId - 1),
            XmlDoc);
        XmlDoc.GetRoot(ShapeRootElement);
        ShapeRootElement.GetChildElements().Get(1, XmlNode);

        DrawinigDoc.GetRoot(DrawinigRootElement);
        DrawinigRootElement.Add(XmlNode);
    end;

    local procedure CreateCommentVmlAnchor(ColId: Integer; RowId: Integer): Text
    begin
        exit(StrSubstNo('%1,15,%2,10,%3,31,%4,9', ColId, RowId - 2, ColId + 2, RowId + 5));
    end;

    procedure AddWorksheetToArchive(var ZipArchive: Codeunit "Data Compression"; SheetId: Integer; var ContentTypesDoc: XmlDocument);
    var
        LegacyDrawingXml: XmlElement;
        RelationId: Text;
    begin
        AddContentType(
            ContentTypesDoc,
            StrSubstNo('/xl/worksheets/sheet%1.xml', SheetId),
            'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');

        if DrawinigCreated then begin
            RelationId :=
                AddRelationship(
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
                    StrSubstNo('../drawings/vmlDrawing%1.vml', SheetId));

            AddToArchive(ZipArchive, DrawinigDoc, StrSubstNo('xl\drawings\vmlDrawing%1.vml', SheetId));

            LegacyDrawingXml := GetWorksheetChildElement('legacyDrawing');
            LegacyDrawingXml.SetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', RelationId);
        end;

        if CommentsCreated then begin
            AddRelationship(
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
                StrSubstNo('../comments%1.xml', SheetId));

            AddContentType(
                ContentTypesDoc,
                StrSubstNo('/xl/comments%1.xml', SheetId),
                'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');

            AddToArchive(ZipArchive, CommentsDoc, StrSubstNo('xl\comments%1.xml', SheetId));
        end;

        if RelationshipCreated then
            AddToArchive(ZipArchive, RelationshipsDoc, StrSubstNo('xl\worksheets\_rels\sheet%1.xml.rels', SheetId));

        AddToArchive(ZipArchive, WorksheetDoc, StrSubstNo('xl\worksheets\sheet%1.xml', SheetId));
    end;

    local procedure AddRelationship(Type: Text; Target: Text): Text;
    var
        RelationshipXml: XmlElement;
        RelationshipsRootXml: XmlElement;
        RelationshipId: Text;
    begin
        if not RelationshipCreated then
            CreateRelationship();

        RelationshipId := GetRelId();
        RelationshipsDoc.GetRoot(RelationshipsRootXml);

        RelationshipXml := XmlElement.Create('Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships');
        RelationshipXml.SetAttribute('Id', RelationshipId);
        RelationshipXml.SetAttribute('Type', Type);
        RelationshipXml.SetAttribute('Target', Target);
        RelationshipsRootXml.Add(RelationshipXml);
        exit(RelationshipId);
    end;

    local procedure AddContentType(var ContentTypesDoc: XmlDocument; PartName: Text; ContentType: Text): Text;
    var
        ContentTypesRootXml: XmlElement;
        OverrideXml: XmlElement;
    begin
        ContentTypesDoc.GetRoot(ContentTypesRootXml);

        OverrideXml := XmlElement.Create('Override', 'http://schemas.openxmlformats.org/package/2006/content-types');
        OverrideXml.SetAttribute('PartName', PartName);
        OverrideXml.SetAttribute('ContentType', ContentType);
        ContentTypesRootXml.Add(OverrideXml);
    end;

    local procedure CreateRelationship()
    begin
        XmlDocument.ReadFrom(
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />',
            RelationshipsDoc);
        RelationshipCreated := true;
    end;

    local procedure AddToArchive(var ZipArchive: Codeunit "Data Compression"; XmlDoc: XmlDocument; Path: Text);
    var
        TempBlob: Codeunit "Temp Blob";
        OutStream: OutStream;
        InStream: InStream;
    begin
        Clear(TempBlob);
        TempBlob.CreateOutStream(OutStream);
        XmlDoc.WriteTo(OutStream);
        TempBlob.CreateInStream(InStream);
        ZipArchive.AddEntry(InStream, Path);
    end;

    local procedure GetRelId(): Text
    begin
        rId += 1;
        exit(StrSubstNo('srId%1', rId));
    end;

    procedure SetPaperSize(PaperSize: Enum "Excel Paper Size"; Orientation: Enum "Excel Page Orientation")
    var
        PageSetupXml: XmlElement;
    begin
        if not ExcelBufferHelper.SelectSingleXmlElement(WorksheetDoc, '/x:worksheet/x:pageSetup', PageSetupXml) then
            Error(CouldNotCreateBookErr);
        PageSetupXml.SetAttribute('paperSize', Format(PaperSize.AsInteger(), 0, 9));
        PageSetupXml.SetAttribute('orientation', LowerCase(Orientation.Names.Get(Orientation.Ordinals.IndexOf(Orientation.AsInteger))));
    end;

    procedure SetPageScale(FitToPage: Enum "Excel Fit Page Scale")
    var
        PageSetupXml: XmlElement;
        SheetPrXml: XmlElement;
        PageSetUpPrXml: XmlElement;

    begin
        if not ExcelBufferHelper.SelectSingleXmlElement(WorksheetDoc, '/x:worksheet/x:pageSetup', PageSetupXml) then
            Error(CouldNotCreateBookErr);

        SheetPrXml := GetWorksheetChildElement('sheetPr');

        if not ExcelBufferHelper.SelectSingleXmlElement(WorksheetDoc, '/x:worksheet/x:sheetPr/x:pageSetUpPr', PageSetUpPrXml) then begin
            if FitToPage = FitToPage::"No scale" then
                exit;
            PageSetUpPrXml := XmlElement.Create('pageSetUpPr', ExcelBufferHelper.MainNamespace());
            SheetPrXml.Add(PageSetUpPrXml);
        end;
        PageSetUpPrXml.SetAttribute('fitToPage', Format(FitToPage <> FitToPage::"No scale", 0, 2));


        if FitToPage in [FitToPage::Columns, FitToPage::Rows] then begin
            if FitToPage = FitToPage::Columns then
                PageSetupXml.SetAttribute('fitToHeight', '0');

            if FitToPage = FitToPage::Rows then
                PageSetupXml.SetAttribute('fitToWidth', '0');
        end;
    end;

    procedure SetPageMargins(Left: Decimal; Top: Decimal; Right: Decimal; Bottom: Decimal; Header: Decimal; Footer: Decimal);
    var
        PageMarginsXml: XmlElement;
    begin
        PageMarginsXml := GetWorksheetChildElement('pageMargins');

        PageMarginsXml.SetAttribute('left', Format(Left, 0, 9));
        PageMarginsXml.SetAttribute('right', Format(Right, 0, 9));
        PageMarginsXml.SetAttribute('top', Format(Top, 0, 9));
        PageMarginsXml.SetAttribute('bottom', Format(Bottom, 0, 9));
        PageMarginsXml.SetAttribute('header', Format(Header, 0, 9));
        PageMarginsXml.SetAttribute('footer', Format(Footer, 0, 9));

    end;

    procedure AddFreezePane(FreezePaneRowNo: Integer; FreezePaneColNo: Integer);
    var
        SheetViewsCollectionXml: XmlElement;
        SheetViewXml: XmlElement;
        PaneXml: XmlElement;
        SelectionXml: XmlElement;
        ColumnCode: Code[10];
    begin
        if (FreezePaneColNo < 2) and (FreezePaneRowNo < 2) then
            exit;

        if FreezePaneRowNo < 1 then
            FreezePaneRowNo := 1;
        if FreezePaneColNo < 1 then
            FreezePaneColNo := 1;

        ColumnCode := ExcelBufferHelper.GetColumnCode(FreezePaneColNo);

        SheetViewsCollectionXml := GetWorksheetChildElement('sheetViews');
        if not ExcelBufferHelper.SelectSingleXmlElement(WorksheetDoc, '/x:worksheet/x:sheetViews/x:sheetView', SheetViewXml) then begin
            SheetViewXml := XmlElement.Create('sheetView', ExcelBufferHelper.MainNamespace());
            SheetViewXml.SetAttribute('workbookViewId', '0');
            SheetViewsCollectionXml.Add(SheetViewXml);
        end;

        PaneXml := XmlElement.Create('pane', ExcelBufferHelper.MainNamespace());
        PaneXml.SetAttribute('state', 'frozen');
        PaneXml.SetAttribute('topLeftCell', StrSubstNo('%1%2', ColumnCode, FreezePaneRowNo));
        case true of
            (FreezePaneColNo > 2) and (FreezePaneRowNo > 2):  // in the middle
                PaneXml.SetAttribute('activePane', 'bottomRight');
            (FreezePaneRowNo > 2):  // Top rows                
                PaneXml.SetAttribute('activePane', 'bottomLeft');
            (FreezePaneColNo > 2):  // Top columns
                PaneXml.SetAttribute('activePane', 'topRight');
        end;
        if FreezePaneRowNo > 1 then
            PaneXml.SetAttribute('ySplit', Format(FreezePaneRowNo - 1, 0, 9));
        if FreezePaneColNo > 1 then
            PaneXml.SetAttribute('xSplit', Format(FreezePaneColNo - 1, 0, 9));

        SheetViewXml.Add(PaneXml);

        if FreezePaneColNo > 1 then
            SheetViewXml.Add(ExcelBufferHelper.CreateElementWithAttribute('selection', 'pane', 'topRight'));

        if FreezePaneRowNo > 1 then
            SheetViewXml.Add(ExcelBufferHelper.CreateElementWithAttribute('selection', 'pane', 'bottomLeft'));

        if (FreezePaneRowNo > 1) and (FreezePaneColNo > 1) then
            SheetViewXml.Add(ExcelBufferHelper.CreateElementWithAttribute('selection', 'pane', 'bottomRight'));
    end;

    procedure SetPageHeaderFooterSettings(DifferentOddAndEvenPages: Boolean; DifferentFirstPage: Boolean)
    var
        HeaderFooterXml: XmlElement;
    begin
        HeaderFooterXml := GetWorksheetChildElement('headerFooter');
        HeaderFooterXml.SetAttribute('differentOddEven', Format(DifferentOddAndEvenPages, 0, 2));
        HeaderFooterXml.SetAttribute('differentFirst', Format(DifferentFirstPage, 0, 2));
    end;

    procedure SetPageHeaderFooterSettings(DifferentOddAndEvenPages: Boolean; DifferentFirstPage: Boolean; ScaleWithDoc: Boolean; AlignWithMargins: Boolean)
    var
        HeaderFooterXml: XmlElement;
    begin
        HeaderFooterXml := GetWorksheetChildElement('headerFooter');
        HeaderFooterXml.SetAttribute('differentOddEven', Format(DifferentOddAndEvenPages, 0, 2));
        HeaderFooterXml.SetAttribute('differentFirst', Format(DifferentFirstPage, 0, 2));
        HeaderFooterXml.SetAttribute('scaleWithDoc', Format(ScaleWithDoc, 0, 2));
        HeaderFooterXml.SetAttribute('alignWithMargins', Format(AlignWithMargins, 0, 2));
    end;

    procedure AddPageHeaderFooter(Type: Enum "Excel Page HeaderFooter Type"; Value: Text)
    var
        HeaderFooterXml: XmlElement;
        ChildElementXml: XmlElement;
        ChildElementName: Text;
    begin
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.headerfooter?view=openxml-3.0.1
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.evenheader?view=openxml-3.0.1        

        HeaderFooterXml := GetWorksheetChildElement('headerFooter');
        ChildElementName := Type.Names.Get(Type.Ordinals.IndexOf(Type.AsInteger));
        if not ExcelBufferHelper.SelectSingleXmlElement(WorksheetDoc, '/x:worksheet/x:headerFooter/x:' + ChildElementName, ChildElementXml) then
            ChildElementXml := ExcelBufferHelper.CreateChildElement(Type.Names(), HeaderFooterXml, ChildElementName);
        ChildElementXml.RemoveNodes();
        ChildElementXml.Add(XmlText.Create(Value));
    end;
}
