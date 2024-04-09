namespace ASD.Excel;
codeunit 58003 "Excel Workbook Style"
{
    var
        ExcelBufferStyle: Record "Excel Buffer Extended" temporary;
        ExcelBufferHelper: Codeunit "Excel Buffer Helper";
        StylesDoc: XmlDocument;
        DefaultFontName: Text;
        DefaultFontSize: Integer;

    procedure InitStylesXml()
    begin
        DefaultFontName := 'Calibri';
        DefaultFontSize := 11;
        // create xl/styles.xml
        XmlDocument.ReadFrom(
            StrSubstNo(
                '<x:styleSheet ' +
                    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
                    'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" ' +
                    'mc:Ignorable="x14ac" ' +
                    'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
                    //'<x:numFmts count="0" />' +
                    '<x:fonts count="1" x14ac:knownFonts="1">' +
                        '<x:font>' +
                            '<x:sz val="%1" />' +
                            '<x:name val="%2" />' +
                            '<x:family val="2" />' +
                            '<x:scheme val="minor" />' +
                        '</x:font>' +
                    '</x:fonts>' +
                    '<x:fills count="2">' +
                        '<x:fill>' +
                            '<x:patternFill patternType="none" />' +
                        '</x:fill>' +
                        '<x:fill>' +
                            '<x:patternFill patternType="gray125" />' +
                        '</x:fill>' +
                    '</x:fills>' +
                    '<x:borders count="1">' +
                        '<x:border>' +
                            '<x:left />' +
                            '<x:right />' +
                            '<x:top />' +
                            '<x:bottom />' +
                            '<x:diagonal />' +
                        '</x:border>' +
                    '</x:borders>' +
                    '<x:cellStyleXfs count="1">' +
                        '<x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" />' +
                    '</x:cellStyleXfs>' +
                    '<x:cellXfs count="1">' +
                        '<x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />' +
                    '</x:cellXfs>' +
                    '<x:cellStyles count="1">' +
                        '<x:cellStyle name="Normal" xfId="0" builtinId="0" />' +
                    '</x:cellStyles> ' +
                '</x:styleSheet>',
                DefaultFontSize, DefaultFontName),
            StylesDoc);

        ExcelBufferStyle.Reset();
        ExcelBufferStyle.DeleteAll();

        Clear(ExcelBufferStyle);
        ExcelBufferStyle."Font Name" := DefaultFontName;
        ExcelBufferStyle."Font Size" := DefaultFontSize;
        ExcelBufferStyle.Insert();

    end;

    procedure GetXmlDocument(): XmlDocument
    begin
        exit(StylesDoc);
    end;

    procedure FindStyleId(ExcelBuffer: Record "Excel Buffer Extended"): Integer
    begin

        if ExcelBuffer."Cell Type" = ExcelBuffer."Cell Type"::Number then
            if ExcelBuffer.NumberFormat = 'General' then
                ExcelBuffer.NumberFormat := '';

        if ExcelBuffer."Cell Type" = ExcelBuffer."Cell Type"::Text then
            if ExcelBuffer.NumberFormat = '@' then
                ExcelBuffer.NumberFormat := '';

        if ExcelBuffer."Cell Type" = ExcelBuffer."Cell Type"::Date then
            if ExcelBuffer.NumberFormat = 'd/m/yyyy' then
                ExcelBuffer.NumberFormat := '';

        if ExcelBuffer."Cell Type" = ExcelBuffer."Cell Type"::Time then
            if ExcelBuffer.NumberFormat = 'H:mm:ss' then
                ExcelBuffer.NumberFormat := '';

        if ExcelBuffer."Font Name" = '' then
            ExcelBuffer."Font Name" := DefaultFontName;

        if ExcelBuffer."Font Size" = 0 then
            ExcelBuffer."Font Size" := DefaultFontSize;

        if (ExcelBuffer."Background Color" <> '') and (ExcelBuffer."Patern Style" = ExcelBuffer."Patern Style"::None) then
            ExcelBuffer."Patern Style" := ExcelBuffer."Patern Style"::Solid;

        if ExcelBuffer."Patern Style" > ExcelBuffer."Patern Style"::None then
            ExcelBuffer."Shading Style" := ExcelBuffer."Shading Style"::None;

        if ExcelBuffer."Shading Style" = ExcelBuffer."Shading Style"::None then begin
            ExcelBuffer."Gradient Color 1" := '';
            ExcelBuffer."Gradient Color 2" := '';
        end;

        if ExcelBuffer."Border Style" <> ExcelBuffer."Border Style"::None then begin
            if ExcelBuffer."Left Border Style" = ExcelBuffer."Left Border Style"::None then
                ExcelBuffer."Left Border Style" := ExcelBuffer."Border Style";
            if ExcelBuffer."Right Border Style" = ExcelBuffer."Right Border Style"::None then
                ExcelBuffer."Right Border Style" := ExcelBuffer."Border Style";
            if ExcelBuffer."Top Border Style" = ExcelBuffer."Top Border Style"::None then
                ExcelBuffer."Top Border Style" := ExcelBuffer."Border Style";
            if ExcelBuffer."Bottom Border Style" = ExcelBuffer."Bottom Border Style"::None then
                ExcelBuffer."Bottom Border Style" := ExcelBuffer."Border Style";

            if ExcelBuffer."Border Color" <> '' then begin
                if ExcelBuffer."Left Border Color" = '' then
                    ExcelBuffer."Left Border Color" := ExcelBuffer."Border Color";
                if ExcelBuffer."Right Border Color" = '' then
                    ExcelBuffer."Right Border Color" := ExcelBuffer."Border Color";
                if ExcelBuffer."Top Border Color" = '' then
                    ExcelBuffer."Top Border Color" := ExcelBuffer."Border Color";
                if ExcelBuffer."Bottom Border Color" = '' then
                    ExcelBuffer."Bottom Border Color" := ExcelBuffer."Border Color";
            end;
        end;

        if not (ExcelBuffer."Horizontal Alignment" in
                [ExcelBuffer."Horizontal Alignment"::Left,
                 ExcelBuffer."Horizontal Alignment"::Right,
                 ExcelBuffer."Horizontal Alignment"::Distributed])
      then
            ExcelBuffer.Indent := 0;

        if not (ExcelBuffer."Text Rotation" IN [0 .. 180]) then
            ExcelBuffer."Text Rotation" := 0;

        if ExcelBuffer."Wrap Text" then
            ExcelBuffer."Shrink To Fit" := FALSE;

        ExcelBufferStyle.Reset();
        ExcelBufferStyle.SetRange(Bold, ExcelBuffer.Bold);
        ExcelBufferStyle.SetRange(Italic, ExcelBuffer.Italic);
        if ExcelBuffer."Double Underline" then
            ExcelBufferStyle.SetRange("Double Underline", ExcelBuffer."Double Underline")
        else
            ExcelBufferStyle.SetRange(Underline, ExcelBuffer.Underline);
        ExcelBufferStyle.SetRange(NumberFormat, ExcelBuffer.NumberFormat);
        ExcelBufferStyle.SetRange("Cell Type", ExcelBuffer."Cell Type");

        ExcelBufferStyle.SetRange("Font Name", ExcelBuffer."Font Name");
        ExcelBufferStyle.SetRange("Font Color", ExcelBuffer."Font Color");
        ExcelBufferStyle.SetRange("Font Size", ExcelBuffer."Font Size");
        ExcelBufferStyle.SetRange(Strikethrough, ExcelBuffer.Strikethrough);
        ExcelBufferStyle.SetRange("VertAlign Effect", ExcelBuffer."VertAlign Effect");

        ExcelBufferStyle.SetRange("Background Color", ExcelBuffer."Background Color");
        ExcelBufferStyle.SetRange("Patern Style", ExcelBuffer."Patern Style");
        ExcelBufferStyle.SetRange("Patern Color", ExcelBuffer."Patern Color");
        ExcelBufferStyle.SetRange("Shading Style", ExcelBuffer."Shading Style");
        ExcelBufferStyle.SetRange("Gradient Color 1", ExcelBuffer."Gradient Color 1");
        ExcelBufferStyle.SetRange("Gradient Color 2", ExcelBuffer."Gradient Color 2");

        ExcelBufferStyle.SetRange("Left Border Style", ExcelBuffer."Left Border Style");
        ExcelBufferStyle.SetRange("Left Border Color", ExcelBuffer."Left Border Color");
        ExcelBufferStyle.SetRange("Right Border Style", ExcelBuffer."Right Border Style");
        ExcelBufferStyle.SetRange("Right Border Color", ExcelBuffer."Right Border Color");
        ExcelBufferStyle.SetRange("Top Border Style", ExcelBuffer."Top Border Style");
        ExcelBufferStyle.SetRange("Top Border Color", ExcelBuffer."Top Border Color");
        ExcelBufferStyle.SetRange("Bottom Border Style", ExcelBuffer."Bottom Border Style");
        ExcelBufferStyle.SetRange("Bottom Border Color", ExcelBuffer."Bottom Border Color");
        ExcelBufferStyle.SetRange("Diagonal Border Style", ExcelBuffer."Diagonal Border Style");
        ExcelBufferStyle.SetRange("Diagonal Border Color", ExcelBuffer."Diagonal Border Color");
        ExcelBufferStyle.SetRange("Diagonal Border Type", ExcelBuffer."Diagonal Border Type");

        ExcelBufferStyle.SetRange("Horizontal Alignment", ExcelBuffer."Horizontal Alignment");
        ExcelBufferStyle.SetRange("Vertical Alignment", ExcelBuffer."Vertical Alignment");
        ExcelBufferStyle.SetRange(Indent, ExcelBuffer.Indent);
        ExcelBufferStyle.SetRange("Reading Order", ExcelBuffer."Reading Order");
        ExcelBufferStyle.SetRange("Relative Indent", ExcelBuffer."Relative Indent");
        ExcelBufferStyle.SetRange("Shrink To Fit", ExcelBuffer."Shrink To Fit");
        ExcelBufferStyle.SetRange("Text Rotation", ExcelBuffer."Text Rotation");
        ExcelBufferStyle.SetRange("Wrap Text", ExcelBuffer."Wrap Text");
        ExcelBufferStyle.SetRange("Justify Last Line", ExcelBuffer."Justify Last Line");


        if ExcelBufferStyle.FindFirst() then
            exit(ExcelBufferStyle."Row No.");

        exit(
            CreateNewStyle(ExcelBuffer));
    end;

    local procedure CreateNewStyle(ExcelBuffer: Record "Excel Buffer Extended"): Integer
    var
        CellXfsXml: XmlElement;
        StyleFormatXml: XmlElement;
        XmlAttr: XmlAttribute;
        StyleId: Integer;
        NumFormatId: Integer;
        FontId: Integer;
        FillId: Integer;
        BorderId: Integer;
    begin
        if not ExcelBufferHelper.SelectSingleXMLElement(StylesDoc, '/x:styleSheet/x:cellXfs', CellXfsXml) then
            exit(0);

        if not CellXfsXml.Attributes.Get('count', XmlAttr) then
            exit(0);

        Evaluate(StyleId, XmlAttr.Value);

        NumFormatId := GetNumberFormat(ExcelBuffer);
        FontId := GetFontId(ExcelBuffer);
        FillId := GetFillId(ExcelBuffer);
        BorderId := GetBorderId(ExcelBuffer);

        StyleFormatXml := XmlElement.Create('xf', ExcelBufferHelper.MainNamespace);
        StyleFormatXml.SetAttribute('numFmtId', Format(NumFormatId, 0, 9));
        StyleFormatXml.SetAttribute('fontId', Format(FontId, 0, 9));
        StyleFormatXml.SetAttribute('fillId', Format(FillId, 0, 9));
        StyleFormatXml.SetAttribute('borderId', Format(BorderId, 0, 9));
        StyleFormatXml.SetAttribute('xfId', '0');
        if NumFormatId <> 0 then
            StyleFormatXml.SetAttribute('applyNumberFormat', '1');
        if FontId <> 0 then
            StyleFormatXml.SetAttribute('applyFont', '1');
        if FillId <> 0 then
            StyleFormatXml.SetAttribute('applyFill', '1');
        if BorderId <> 0 then
            StyleFormatXml.SetAttribute('applyBorder', '1');

        AddAlignment(StyleFormatXml, ExcelBuffer);

        CellXfsXml.Add(StyleFormatXml);
        CellXfsXml.SetAttribute('count', Format(StyleId + 1, 0, 9));

        Clear(ExcelBufferStyle);
        ExcelBufferStyle."Row No." := StyleId;
        ExcelBufferStyle.Bold := ExcelBuffer.Bold;
        ExcelBufferStyle.Italic := ExcelBuffer.Italic;
        ExcelBufferStyle.Underline := ExcelBuffer.Underline;
        ExcelBufferStyle.NumberFormat := ExcelBuffer.NumberFormat;
        ExcelBufferStyle."Cell Type" := ExcelBuffer."Cell Type";
        ExcelBufferStyle."Double Underline" := ExcelBuffer."Double Underline";

        ExcelBufferStyle."Font Name" := ExcelBuffer."Font Name";
        ExcelBufferStyle."Font Color" := ExcelBuffer."Font Color";
        ExcelBufferStyle."Font Size" := ExcelBuffer."Font Size";
        ExcelBufferStyle.Strikethrough := ExcelBuffer.Strikethrough;
        ExcelBufferStyle."VertAlign Effect" := ExcelBuffer."VertAlign Effect";

        ExcelBufferStyle."Background Color" := ExcelBuffer."Background Color";
        ExcelBufferStyle."Patern Style" := ExcelBuffer."Patern Style";
        ExcelBufferStyle."Patern Color" := ExcelBuffer."Patern Color";
        ExcelBufferStyle."Shading Style" := ExcelBuffer."Shading Style";
        ExcelBufferStyle."Gradient Color 1" := ExcelBuffer."Gradient Color 1";
        ExcelBufferStyle."Gradient Color 2" := ExcelBuffer."Gradient Color 2";

        ExcelBufferStyle."Left Border Style" := ExcelBuffer."Left Border Style";
        ExcelBufferStyle."Left Border Color" := ExcelBuffer."Left Border Color";
        ExcelBufferStyle."Right Border Style" := ExcelBuffer."Right Border Style";
        ExcelBufferStyle."Right Border Color" := ExcelBuffer."Right Border Color";
        ExcelBufferStyle."Top Border Style" := ExcelBuffer."Top Border Style";
        ExcelBufferStyle."Top Border Color" := ExcelBuffer."Top Border Color";
        ExcelBufferStyle."Bottom Border Style" := ExcelBuffer."Bottom Border Style";
        ExcelBufferStyle."Bottom Border Color" := ExcelBuffer."Bottom Border Color";
        ExcelBufferStyle."Diagonal Border Style" := ExcelBuffer."Diagonal Border Style";
        ExcelBufferStyle."Diagonal Border Color" := ExcelBuffer."Diagonal Border Color";
        ExcelBufferStyle."Diagonal Border Type" := ExcelBuffer."Diagonal Border Type";

        ExcelBufferStyle."Horizontal Alignment" := ExcelBuffer."Horizontal Alignment";
        ExcelBufferStyle."Vertical Alignment" := ExcelBuffer."Vertical Alignment";
        ExcelBufferStyle.Indent := ExcelBuffer.Indent;
        ExcelBufferStyle."Reading Order" := ExcelBuffer."Reading Order";
        ExcelBufferStyle."Relative Indent" := ExcelBuffer."Relative Indent";
        ExcelBufferStyle."Shrink To Fit" := ExcelBuffer."Shrink To Fit";
        ExcelBufferStyle."Text Rotation" := ExcelBuffer."Text Rotation";
        ExcelBufferStyle."Wrap Text" := ExcelBuffer."Wrap Text";
        ExcelBufferStyle."Justify Last Line" := ExcelBuffer."Justify Last Line";

        ExcelBufferStyle.Insert();
        Exit(StyleId);

    end;

    local procedure GetNumberFormat(ExcelBuffer: record "Excel Buffer Extended"): Integer
    var
        NumFormatText: Text;
    begin
        NumFormatText := ExcelBuffer.NumberFormat;
        if NumFormatText = '' then
            Case ExcelBuffer."Cell Type" of
                ExcelBuffer."Cell Type"::Number:
                    NumFormatText := 'General';
                ExcelBuffer."Cell Type"::Text:
                    NumFormatText := '@';
                ExcelBuffer."Cell Type"::Date:
                    NumFormatText := 'd/m/yyyy';
                ExcelBuffer."Cell Type"::Time:
                    NumFormatText := 'H:mm:ss';
            end;

        case NumFormatText of
            'General', '':
                Exit(0);
            '0':
                Exit(1);
            '0.00':
                Exit(2);
            '#,##0':
                Exit(3);
            '#,##0.00':
                Exit(4);
            '0%':
                Exit(9);
            '0.00%':
                Exit(10);
            '0.00E+00':
                Exit(11);
            '# ?/?':
                Exit(12);
            '# ??/??':
                Exit(13);
            'd/m/yyyy':
                Exit(14);
            'd-mmm-yy':
                Exit(15);
            'd-mmm':
                Exit(16);
            'mmm-yy':
                Exit(17);
            'h:mm tt':
                Exit(18);
            'h:mm:ss tt':
                Exit(19);
            'H:mm':
                Exit(20);
            'H:mm:ss':
                Exit(21);
            'm/d/yyyy H:mm':
                Exit(22);
            '#,##0 );(#,##0)':
                Exit(37);
            '#,##0 );[Red](#,##0)':
                Exit(38);
            '#,##0.00);(#,##0.00)':
                Exit(39);
            '#,##0.00);[Red](#,##0.00)':
                Exit(40);
            'mm:ss':
                Exit(45);
            '[h]:mm:ss':
                Exit(46);
            'mmss.0':
                Exit(47);
            '##0.0E+0':
                Exit(48);
            '@':
                Exit(49);
            else begin
                Exit(GetCustomNumberFormat(NumFormatText));
            end;
        end;
        /*
            ID  Format Code
            0   General
            1   0
            2   0.00
            3   #,##0
            4   #,##0.00
            9   0%
            10  0.00%
            11  0.00E+00
            12  # ?/?
            13  # ??/??
            14  d/m/yyyy
            15  d-mmm-yy
            16  d-mmm
            17  mmm-yy
            18  h:mm tt
            19  h:mm:ss tt
            20  H:mm
            21  H:mm:ss
            22  m/d/yyyy H:mm
            37  #,##0 ;(#,##0)
            38  #,##0 ;[Red](#,##0)
            39  #,##0.00;(#,##0.00)
            40  #,##0.00;[Red](#,##0.00)
            45  mm:ss
            46  [h]:mm:ss
            47  mmss.0
            48  ##0.0E+0
            49  @
            */
    end;

    local procedure GetCustomNumberFormat(NumFormatText: Text): Integer
    var
        numFmtsCollectionXml: XmlElement;
        numFmtXml: XmlElement;
        RootElement: XmlElement;
        XmlAttr: XmlAttribute;
        XmlNode: XmlNode;
        numFmtId: Integer;
        xPath: Text;
        CountValue: Integer;
    begin
        if not ExcelBufferHelper.SelectSingleXMLElement(StylesDoc, '/x:styleSheet/x:numFmts', numFmtsCollectionXml) then begin
            numFmtsCollectionXml := XmlElement.Create('numFmts', ExcelBufferHelper.MainNamespace);
            numFmtsCollectionXml.SetAttribute('count', '0');
            StylesDoc.GetRoot(RootElement);
            RootElement.AddFirst(numFmtsCollectionXml);
        end else begin
            xPath := StrSubstNo('/x:styleSheet/x:numFmts/x:numFmt[@formatCode=''%1'']', NumFormatText);
            if not ExcelBufferHelper.SelectSingleXMLElement(StylesDoc, xPath, numFmtXml) then begin
                if numFmtXml.Attributes().Get('numFmtId', XmlAttr) then
                    if Evaluate(numFmtId, XmlAttr.Value) then
                        exit(numFmtId);
            end;
        end;
        if numFmtsCollectionXml.GetChildElements().Count > 0 then begin
            numFmtsCollectionXml.GetChildElements().Get(numFmtsCollectionXml.GetChildElements().Count - 1, XmlNode);
            numFmtXml := XmlNode.AsXmlElement();
            if numFmtXml.Attributes().Get('numFmtId', XmlAttr) then
                Evaluate(numFmtId, XmlAttr.Value);
        end;

        if numFmtId < 164 then
            numFmtId := 164
        else
            numFmtId += 1;

        numFmtXml := XmlElement.Create('numFmt', ExcelBufferHelper.MainNamespace);
        numFmtXml.SetAttribute('numFmtId', Format(numFmtId, 0, 9));
        numFmtXml.SetAttribute('formatCode', NumFormatText);
        numFmtsCollectionXml.Add(numFmtXml);
        if numFmtsCollectionXml.Attributes().Get('count', XmlAttr) then begin
            Evaluate(CountValue, XmlAttr.Value, 9);
            numFmtsCollectionXml.SetAttribute('count', format(CountValue + 1, 0, 9));
        end;
        exit(numFmtId);
    end;

    local procedure GetFontId(ExcelBuffer: Record "Excel Buffer Extended"): Integer
    var
        XMLNamespaceManager: XmlNamespaceManager;
        XmlNode: XmlNode;
        XmlNodeFontName: XmlNode;
        FontsCollectionXml: XmlElement;
        NewFontXml: XmlElement;
        XmlAttr: XmlAttribute;
        FontId: Integer;
        NoOfKnownFonts: Integer;
        IsNewFontName: Boolean;
    begin
        if not ExcelBufferHelper.SelectSingleXMLElement(StylesDoc, '/x:styleSheet/x:fonts', FontsCollectionXml) then
            exit(0);

        if ExcelBuffer."Font Name" = '' then
            ExcelBuffer."Font Name" := DefaultFontName;
        if ExcelBuffer."Font Size" = 0 then
            ExcelBuffer."Font Size" := DefaultFontSize;

        NewFontXml := XmlElement.Create('font', ExcelBufferHelper.MainNamespace);

        if ExcelBuffer.Bold then
            NewFontXml.Add(XmlElement.Create('b', ExcelBufferHelper.MainNamespace));
        if ExcelBuffer.Italic then
            NewFontXml.Add(XmlElement.Create('i', ExcelBufferHelper.MainNamespace));
        Case true of
            ExcelBuffer."Double Underline":
                NewFontXml.Add(ExcelBufferHelper.CreateElementWithAttribute('u', 'val', 'double'));
            //<u val="singleAccounting"/>
            //<u val="doubleAccounting"/>
            ExcelBuffer.Underline:
                NewFontXml.Add(XmlElement.Create('u', ExcelBufferHelper.MainNamespace));
        End;
        if ExcelBuffer.Strikethrough then
            NewFontXml.Add(XmlElement.Create('strike', ExcelBufferHelper.MainNamespace));

        Case ExcelBuffer."VertAlign Effect" of
            ExcelBuffer."VertAlign Effect"::Superscript:
                NewFontXml.Add(ExcelBufferHelper.CreateElementWithAttribute('vertAlign', 'val', 'superscript'));
            ExcelBuffer."VertAlign Effect"::Subscript:
                NewFontXml.Add(ExcelBufferHelper.CreateElementWithAttribute('vertAlign', 'val', 'subscript'));
        End;

        NewFontXml.Add(ExcelBufferHelper.CreateElementWithAttribute('sz', 'val', Format(ExcelBuffer."Font Size", 0, 9)));

        if ExcelBuffer."Font Color" <> '' then
            NewFontXml.Add(ExcelBufferHelper.CreateElementWithAttribute('color', 'rgb', ExcelBuffer."Font Color"));

        NewFontXml.Add(ExcelBufferHelper.CreateElementWithAttribute('name', 'val', ExcelBuffer."Font Name"));
        NewFontXml.Add(ExcelBufferHelper.CreateElementWithAttribute('family', 'val', '2'));

        if ExcelBuffer."Font Name" = DefaultFontName then
            NewFontXml.Add(ExcelBufferHelper.CreateElementWithAttribute('scheme', 'val', 'minor'));

        FontId := 0;
        IsNewFontName := true;
        ExcelBufferHelper.InitXMLNamespaceMngr(StylesDoc, XMLNamespaceManager);
        foreach XmlNode in FontsCollectionXml.GetChildElements() do begin
            if ExcelBufferHelper.AreXMLElementsTheSame(XmlNode.AsXmlElement(), NewFontXml) then
                exit(FontId);
            if XmlNode.AsXmlElement().SelectSingleNode('//x:name', XMLNamespaceManager, XmlNodeFontName) then begin
                if XmlNodeFontName.AsXmlElement().Attributes().Get('val', XmlAttr) then
                    if XmlAttr.Value = ExcelBuffer."Font Name" then
                        IsNewFontName := false;
            end;
            FontId += 1;
        end;

        FontsCollectionXml.Add(NewFontXml);
        FontsCollectionXml.SetAttribute('count', Format(FontId + 1, 0, 9));
        if IsNewFontName then
            if FontsCollectionXml.Attributes().Get('knownFonts', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac', XmlAttr) then begin
                Evaluate(NoOfKnownFonts, XmlAttr.Value, 9);
                FontsCollectionXml.SetAttribute('knownFonts', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac', format(NoOfKnownFonts + 1, 0, 9));
            end;
        exit(FontId);
    end;

    local procedure GetFillId(ExcelBuffer: record "Excel Buffer Extended"): Integer
    var
        FillsCollectionXml: XmlElement;
        NewFillXml: XmlElement;
        NewChildFillXml: XmlElement;
        XmlAttr: XmlAttribute;
        FillId: Integer;
        PaternStyleValue: Text;
    begin
        if (ExcelBuffer."Background Color" = '') and
            (ExcelBuffer."Patern Style" <= ExcelBuffer."Patern Style"::Solid) and
            ((ExcelBuffer."Shading Style" = ExcelBuffer."Shading Style"::None) or (ExcelBuffer."Gradient Color 1" = ExcelBuffer."Gradient Color 2"))
        then
            exit(0);

        if (ExcelBuffer."Background Color" <> '') and (ExcelBuffer."Patern Style" = ExcelBuffer."Patern Style"::None) then
            ExcelBuffer."Patern Style" := ExcelBuffer."Patern Style"::Solid;

        if not ExcelBufferHelper.SelectSingleXMLElement(StylesDoc, '/x:styleSheet/x:fills', FillsCollectionXml) then
            exit(0);


        NewFillXml := XmlElement.Create('fill', ExcelBufferHelper.MainNamespace);
        if ExcelBuffer."Patern Style" > ExcelBuffer."Patern Style"::None then begin
            NewChildFillXml := XmlElement.Create('patternFill', ExcelBufferHelper.MainNamespace);
            Case ExcelBuffer."Patern Style" of
                ExcelBuffer."Patern Style"::Solid:
                    PaternStyleValue := 'solid';
                ExcelBuffer."Patern Style"::"75% Gray":
                    PaternStyleValue := 'darkGray';
                ExcelBuffer."Patern Style"::"50% Gray":
                    PaternStyleValue := 'mediumGray';
                ExcelBuffer."Patern Style"::"25% Gray":
                    PaternStyleValue := 'lightGray';
                ExcelBuffer."Patern Style"::"12.5% Gray":
                    PaternStyleValue := 'gray125';
                ExcelBuffer."Patern Style"::"6.25% Gray":
                    PaternStyleValue := 'gray0625';
                ExcelBuffer."Patern Style"::"Horizontal Stripe":
                    PaternStyleValue := 'darkHorizontal';
                ExcelBuffer."Patern Style"::"Vertical Stripe":
                    PaternStyleValue := 'darkVertical';
                ExcelBuffer."Patern Style"::"Reverse Diagonal Stripe":
                    PaternStyleValue := 'darkDown';
                ExcelBuffer."Patern Style"::"Diagonal Stripe":
                    PaternStyleValue := 'darkUp';
                ExcelBuffer."Patern Style"::"Diagonal Crosshatch":
                    PaternStyleValue := 'darkGrid';
                ExcelBuffer."Patern Style"::"Thick Diagonal Crosshatch":
                    PaternStyleValue := 'darkTrellis';
                ExcelBuffer."Patern Style"::"Thin Horizontal Stripe":
                    PaternStyleValue := 'lightHorizontal';
                ExcelBuffer."Patern Style"::"Thin Vertical Stripe":
                    PaternStyleValue := 'lightVertical';
                ExcelBuffer."Patern Style"::"Thin Reverse Diagonal Stripe":
                    PaternStyleValue := 'lightDown';
                ExcelBuffer."Patern Style"::"Thin Diagonal Stripe":
                    PaternStyleValue := 'lightUp';
                ExcelBuffer."Patern Style"::"Thin Horizontal Crosshatch":
                    PaternStyleValue := 'lightGrid';
                ExcelBuffer."Patern Style"::"Thin Diagonal Crosshatch":
                    PaternStyleValue := 'lightTrellis';
            End;
            if PaternStyleValue <> '' then
                NewChildFillXml.SetAttribute('patternType', PaternStyleValue);
            if ExcelBuffer."Background Color" <> '' then
                NewChildFillXml.Add(ExcelBufferHelper.CreateElementWithAttribute('fgColor', 'rgb', ExcelBuffer."Background Color"));

            if ExcelBuffer."Patern Color" <> '' then
                NewChildFillXml.Add(ExcelBufferHelper.CreateElementWithAttribute('bgColor', 'rgb', ExcelBuffer."Patern Color"))
            else
                if (ExcelBuffer."Background Color" <> '') and (ExcelBuffer."Patern Style" = ExcelBuffer."Patern Style"::Solid) then
                    NewChildFillXml.Add(ExcelBufferHelper.CreateElementWithAttribute('bgColor', 'auto', '1'));
            NewFillXml.Add(NewChildFillXml);
        end else
            if ExcelBuffer."Shading Style" <> ExcelBuffer."Shading Style"::None then begin
                NewChildFillXml := XmlElement.Create('gradientFill', ExcelBufferHelper.MainNamespace);
                case ExcelBuffer."Shading Style" of
                    ExcelBuffer."Shading Style"::Horizontal:
                        begin
                            NewChildFillXml.SetAttribute('degree', '90');
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 2"));
                        end;
                    ExcelBuffer."Shading Style"::"Horizontal Middle":
                        begin
                            NewChildFillXml.SetAttribute('degree', '90');
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(0.5, ExcelBuffer."Gradient Color 2"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 1"));
                        end;
                    ExcelBuffer."Shading Style"::Vertical:
                        begin
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 2"));
                        end;
                    ExcelBuffer."Shading Style"::"Vertical Middle":
                        begin
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(0.5, ExcelBuffer."Gradient Color 2"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 1"));
                        end;
                    ExcelBuffer."Shading Style"::"Diagonal Up":
                        begin
                            NewChildFillXml.SetAttribute('degree', '45');
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 2"));
                        end;
                    ExcelBuffer."Shading Style"::"Diagonal Up Middle":
                        begin
                            NewChildFillXml.SetAttribute('degree', '45');
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(0.5, ExcelBuffer."Gradient Color 2"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 1"));
                        end;
                    ExcelBuffer."Shading Style"::"Diagonal Down":
                        begin
                            NewChildFillXml.SetAttribute('degree', '135');
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 2"));
                        end;
                    ExcelBuffer."Shading Style"::"Diagonal Down Middle":
                        begin
                            NewChildFillXml.SetAttribute('degree', '135');
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(0.5, ExcelBuffer."Gradient Color 2"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 1"));
                        end;
                    ExcelBuffer."Shading Style"::"From Left Top Corner":
                        begin
                            NewChildFillXml.SetAttribute('type', 'path');
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 2"));
                        end;
                    ExcelBuffer."Shading Style"::"From Right Top Corner":
                        begin
                            NewChildFillXml.SetAttribute('type', 'path');
                            NewChildFillXml.SetAttribute('left', '1');
                            NewChildFillXml.SetAttribute('right', '1');
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 2"));
                        end;
                    ExcelBuffer."Shading Style"::"From Left Bottom Corner":
                        begin
                            NewChildFillXml.SetAttribute('type', 'path');
                            NewChildFillXml.SetAttribute('top', '1');
                            NewChildFillXml.SetAttribute('bottom', '1');
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 2"));
                        end;
                    ExcelBuffer."Shading Style"::"From Right Bottom Corner":
                        begin
                            NewChildFillXml.SetAttribute('type', 'path');
                            NewChildFillXml.SetAttribute('left', '1');
                            NewChildFillXml.SetAttribute('right', '1');
                            NewChildFillXml.SetAttribute('top', '1');
                            NewChildFillXml.SetAttribute('bottom', '1');
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 2"));
                        end;
                    ExcelBuffer."Shading Style"::"From Center":
                        begin
                            NewChildFillXml.SetAttribute('type', 'path');
                            NewChildFillXml.SetAttribute('left', '0.5');
                            NewChildFillXml.SetAttribute('right', '0.5');
                            NewChildFillXml.SetAttribute('top', '0.5');
                            NewChildFillXml.SetAttribute('bottom', '0.5');
                            NewChildFillXml.Add(CreateGradientPosition(0, ExcelBuffer."Gradient Color 1"));
                            NewChildFillXml.Add(CreateGradientPosition(1, ExcelBuffer."Gradient Color 2"));
                        end;
                end;
                NewFillXml.Add(NewChildFillXml);
            end;

        if ExcelBufferHelper.FindElementInCollection(FillsCollectionXml, NewFillXml, FillId) then
            exit(FillId);

        FillsCollectionXml.Add(NewFillXml);
        FillsCollectionXml.SetAttribute('count', Format(FillId + 1, 0, 9));
        exit(FillId);

    end;

    local procedure CreateGradientPosition(Postion: Decimal; Color: Text): XmlElement
    var
        StopPositionXml: XmlElement;
        ColorXml: XmlElement;
    begin
        StopPositionXml := XmlElement.Create('stop', ExcelBufferHelper.MainNamespace);
        StopPositionXml.SetAttribute('position', Format(Postion, 0, 9));
        ColorXml := XmlElement.Create('color', ExcelBufferHelper.MainNamespace);
        ColorXml.SetAttribute('rgb', Color);
        StopPositionXml.Add(ColorXml);
        Exit(StopPositionXml);
    end;

    local procedure GetBorderId(ExcelBuffer: record "Excel Buffer Extended"): Integer
    var
        BordersCollectionXml: XmlElement;
        NewBorderXml: XmlElement;
        XmlAttr: XmlAttribute;
        BorderId: Integer;
    begin

        if (ExcelBuffer."Left Border Style" = ExcelBuffer."Left Border Style"::None) and
            (ExcelBuffer."Right Border Style" = ExcelBuffer."Right Border Style"::None) and
            (ExcelBuffer."Top Border Style" = ExcelBuffer."Top Border Style"::None) and
            (ExcelBuffer."Bottom Border Style" = ExcelBuffer."Bottom Border Style"::None) and
            (ExcelBuffer."Diagonal Border Style" = ExcelBuffer."Diagonal Border Style"::None)
        then
            exit(0);

        if not ExcelBufferHelper.SelectSingleXMLElement(StylesDoc, '/x:styleSheet/x:borders', BordersCollectionXml) then
            exit(0);

        NewBorderXml := XmlElement.Create('border', ExcelBufferHelper.MainNamespace);

        NewBorderXml.Add(CreateBorderElement('left', ExcelBuffer."Left Border Style", ExcelBuffer."Left Border Color"));
        NewBorderXml.Add(CreateBorderElement('right', ExcelBuffer."Right Border Style", ExcelBuffer."Right Border Color"));
        NewBorderXml.Add(CreateBorderElement('top', ExcelBuffer."Top Border Style", ExcelBuffer."Top Border Color"));
        NewBorderXml.Add(CreateBorderElement('bottom', ExcelBuffer."Bottom Border Style", ExcelBuffer."Bottom Border Color"));

        NewBorderXml.Add(CreateBorderElement('diagonal', ExcelBuffer."Diagonal Border Style", ExcelBuffer."Diagonal Border Color"));
        if ExcelBuffer."Diagonal Border Style" <> ExcelBuffer."Diagonal Border Style"::None then begin
            if ExcelBuffer."Diagonal Border Type" in [ExcelBuffer."Diagonal Border Type"::Up, ExcelBuffer."Diagonal Border Type"::"Up and Down"] then
                NewBorderXml.SetAttribute('diagonalUp', '1');
            if ExcelBuffer."Diagonal Border Type" in [ExcelBuffer."Diagonal Border Type"::Down, ExcelBuffer."Diagonal Border Type"::"Up and Down"] then
                NewBorderXml.SetAttribute('diagonalDown', '1');
        end;

        if ExcelBufferHelper.FindElementInCollection(BordersCollectionXml, NewBorderXml, BorderId) then
            exit(BorderId);

        BordersCollectionXml.Add(NewBorderXml);
        BordersCollectionXml.SetAttribute('count', Format(BorderId + 1, 0, 9));
        exit(BorderId);

    end;

    local procedure CreateBorderElement(ElementName: text; BorderStyle: Enum "Excel Buffer Border Style"; BorderColor: Text): XmlElement
    var
        BorderXml: XmlElement;
        ColorXml: XmlElement;
        BorderStyleName: Text;
    begin
        BorderXml := XmlElement.Create(ElementName, ExcelBufferHelper.MainNamespace);
        if BorderStyle = BorderStyle::None then
            exit(BorderXml);

        BorderStyleName := BorderStyle.Names.Get(BorderStyle.Ordinals.IndexOf(BorderStyle.AsInteger));
        BorderStyleName := BorderStyleName.Replace(' ', '');
        BorderStyleName := LowerCase(BorderStyleName[1]) + BorderStyleName.Substring(2);

        BorderXml.SetAttribute('style', BorderStyleName);
        ColorXml := XmlElement.Create('color', ExcelBufferHelper.MainNamespace);
        if BorderColor = '' then
            ColorXml.SetAttribute('auto', '1')
        else
            ColorXml.SetAttribute('rgb', BorderColor);
        BorderXml.Add(ColorXml);

        exit(BorderXml);
    end;

    local procedure AddAlignment(var StyleFormatXml: XmlElement; ExcelBuffer: record "Excel Buffer Extended")
        AlignmentXml: XmlElement;
    begin
        if (ExcelBuffer."Horizontal Alignment" = ExcelBuffer."Horizontal Alignment"::Automatic) and
         (ExcelBuffer."Vertical Alignment" = ExcelBuffer."Vertical Alignment"::Automatic) and
         (ExcelBuffer.Indent = 0) and
         (ExcelBuffer."Reading Order" = ExcelBuffer."Reading Order"::Context) and
         (ExcelBuffer."Relative Indent" = 0) and
         (ExcelBuffer."Shrink To Fit" = false) and
         (ExcelBuffer."Text Rotation" = 0) and
         (ExcelBuffer."Wrap Text" = false) and
         (ExcelBuffer."Justify Last Line" = false)
      //(ExcelBuffer.Unlocked = false) and
      //(ExcelBuffer."Formula Hidden" = false)
      then
            exit;
        AlignmentXml := XmlElement.Create('alignment', ExcelBufferHelper.MainNamespace);
        IF ExcelBuffer."Horizontal Alignment" <> ExcelBuffer."Horizontal Alignment"::Automatic THEN
            AlignmentXml.SetAttribute('horizontal', LowerCase(Format(ExcelBuffer."Horizontal Alignment")));
        IF ExcelBuffer."Vertical Alignment" <> ExcelBuffer."Vertical Alignment"::Automatic THEN
            AlignmentXml.SetAttribute('vertical', LowerCase(Format(ExcelBuffer."Vertical Alignment")));
        IF ExcelBuffer.Indent > 0 THEN
            AlignmentXml.SetAttribute('indent', Format(ExcelBuffer.Indent, 0, 9));

        IF ExcelBuffer."Reading Order" <> ExcelBuffer."Reading Order"::Context THEN
            AlignmentXml.SetAttribute('readingOrder', Format(ExcelBuffer."Reading Order", 0, 9));

        IF ExcelBuffer."Relative Indent" > 0 THEN
            AlignmentXml.SetAttribute('relativeIndent', Format(ExcelBuffer."Relative Indent", 0, 9));

        IF ExcelBuffer."Shrink To Fit" THEN
            AlignmentXml.SetAttribute('shrinkToFit', Format(ExcelBuffer."Shrink To Fit", 0, 9));

        IF ExcelBuffer."Text Rotation" > 0 THEN
            AlignmentXml.SetAttribute('textRotation', Format(ExcelBuffer."Text Rotation", 0, 9));

        IF ExcelBuffer."Wrap Text" THEN
            AlignmentXml.SetAttribute('wrapText', Format(ExcelBuffer."Wrap Text", 0, 9));

        IF ExcelBuffer."Justify Last Line" THEN
            AlignmentXml.SetAttribute('justifyLastLine', Format(ExcelBuffer."Justify Last Line", 0, 9));
        StyleFormatXml.Add(AlignmentXml);
        StyleFormatXml.SetAttribute('applyAlignment', '1');

    end;
}
