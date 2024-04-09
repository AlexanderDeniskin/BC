namespace ASD.Excel;

using System.IO;
using System.Utilities;
codeunit 58001 "Excel Workbook"
{
    var
        WorkbookStyle: Codeunit "Excel Workbook Style";
        ExcelBufferHelper: Codeunit "Excel Buffer Helper";
        ZipArchive: Codeunit "Data Compression";
        ContentTypesDoc: XmlDocument;
        GlobalRelsDoc: XmlDocument;
        WorkbookDoc: XmlDocument;
        WorkbookRelsDoc: XmlDocument;
        rId: Integer;
        CurrSheetId: Integer;
        EnterWkshtNameErr: Label 'You must enter an Excel worksheet name.';
        CouldNotCreateBookErr: Label 'Could not create the Excel workbook.';
        CreatingExcelWorksheetTxt: Label 'Creating Excel worksheet...\\', Comment = '{Locked="Excel"}';
        OperationCanceledErr: Label 'The operation was canceled.';

    procedure CreateBook()
    begin
        rID := 0;
        CurrSheetId := 0;

        InitContentTypeXml();
        WorkbookStyle.InitStylesXml();
        InitGlobalRelationsXml();
        InitWorkbookXml();
        InitWorkbookRelsXml();

        ZipArchive.CreateZipArchive();
    end;

    procedure CreateBook(SheetName: Text; var Worksheet: Codeunit "Excel Worksheet")
    begin
        if SheetName = '' then
            Error(EnterWkshtNameErr);
        CreateBook();
        AddNewSheet(SheetName, Worksheet);
    end;

    procedure AddNewSheet(SheetName: Text; var Worksheet: Codeunit "Excel Worksheet")
    var
        XmlElmnt: XmlElement;
        XmlChildElmnt: XmlElement;
        WkshRelId: Text;
        NewSheetId: Integer;
    begin
        if SheetName = '' then
            Error(EnterWkshtNameErr);

        WkshRelId := GetRelId();
        if not ExcelBufferHelper.SelectSingleXMLElement(WorkbookDoc, '/x:workbook/x:sheets', XmlElmnt) then
            Error(CouldNotCreateBookErr);

        NewSheetId := XmlElmnt.GetChildElements().Count + 1;
        CurrSheetId := NewSheetId;

        XmlChildElmnt := xmlElement.Create('sheet', ExcelBufferHelper.MainNamespace);
        XmlChildElmnt.SetAttribute('name', SheetName);
        XmlChildElmnt.SetAttribute('sheetId', Format(NewSheetId, 0, 9));
        XmlChildElmnt.SetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', WkshRelId);
        XmlChildElmnt.RemoveAttribute('xmlns', ExcelBufferHelper.MainNamespace);
        XmlElmnt.Add(XmlChildElmnt);

        WorkbookRelsDoc.GetRoot(XmlElmnt);

        XmlChildElmnt := xmlElement.Create('Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships');
        XmlChildElmnt.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
        XmlChildElmnt.SetAttribute('Target', StrSubstNo('/xl/worksheets/sheet%1.xml', NewSheetId));
        XmlChildElmnt.SetAttribute('Id', WkshRelId);
        XmlElmnt.Add(XmlChildElmnt);

        Worksheet.InitWorksheet();
        Worksheet.SetWorkbookStyle(WorkbookStyle);
    end;

    procedure SaveWorksheet(Worksheet: Codeunit "Excel Worksheet");
    begin
        Worksheet.AddWorksheetToArchive(ZipArchive, CurrSheetId, ContentTypesDoc);
    end;

    local procedure InitContentTypeXml()
    begin
        // create [Content_Types].xml
        XmlDocument.ReadFrom(
            '<Types ' +
                'xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                '<Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />' +
                '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing" />' +
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />' +
                '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />' +
            '</Types>',
            ContentTypesDoc);

    end;

    local procedure InitGlobalRelationsXml()
    begin
        // create _rels/.rels
        XmlDocument.ReadFrom(
            StrSubstNo(
                '<Relationships ' +
                    'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship ' +
                        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ' +
                        'Target="/xl/workbook.xml" Id="%1" />' +
                '</Relationships>',
                GetRelId()),
            GlobalRelsDoc);
    end;

    local procedure InitWorkbookXml()
    begin
        // create xl\workbook.xml
        XmlDocument.ReadFrom(
            '<x:workbook ' +
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
                'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                '<x:sheets xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />' +
            '</x:workbook>',
            WorkbookDoc);
    end;

    local procedure InitWorkbookRelsXml()
    begin
        // create xl\_rels\workbook.xml.rels
        XmlDocument.ReadFrom(
            StrSubstNo(
                '<Relationships ' +
                    'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="/xl/styles.xml" Id="Re87b5e1e69f14d41" />' +
                '</Relationships>',
                GetRelId()),
            WorkbookRelsDoc);
    end;

    procedure CloseBook()
    begin
        AddToArchive(ContentTypesDoc, '[Content_Types].xml');
        AddToArchive(GlobalRelsDoc, '_rels\.rels');
        AddToArchive(WorkbookStyle.GetXmlDocument(), 'xl\styles.xml');
        AddToArchive(WorkbookDoc, 'xl\workbook.xml');
        AddToArchive(WorkbookRelsDoc, 'xl\_rels\workbook.xml.rels');
    end;

    local procedure AddToArchive(XmlDoc: XmlDocument; Path: Text);
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

    procedure DownloadBook(FileName: Text)
    var
        TempBlob: Codeunit "Temp Blob";
        OutStream: OutStream;
        InStream: InStream;
    begin
        Clear(TempBlob);
        TempBlob.CreateOutStream(OutStream);
        ZipArchive.SaveZipArchive(OutStream);
        TempBlob.CreateInStream(InStream);
        DownloadFromStream(InStream, '', '', '', FileName);
    end;

    procedure SaveToStream(var OutStream: OutStream)
    begin
        ZipArchive.SaveZipArchive(OutStream);
    end;

    local procedure GetRelId(): Text
    begin
        rId += 1;
        Exit(StrSubstNo('rId%1', rId));
    end;

}
