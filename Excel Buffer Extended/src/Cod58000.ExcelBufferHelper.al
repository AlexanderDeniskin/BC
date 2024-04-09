namespace ASD.Excel;
codeunit 58001 "Excel Buffer Helper"
{
    var
        InvalidColumnCodeErr: Label '%1 is invalid Column Code.';
        InvalidCellAddrErr: Label '%1 is invalid cell address.';

    procedure GetColumnCode(ColumnId: Integer) ReturnText: Text[10];
    var
        x: Integer;
        i: Integer;
        y: Integer;
        c: Char;
        t: Text[30];
    begin
        ReturnText := '';

        if ColumnId < 1 then
            exit;

        x := ColumnId;
        while x > 26 do begin
            y := x mod 26;
            if y = 0 then
                y := 26;
            c := 64 + y;
            i := i + 1;
            t[i] := c;
            x := (x - y) div 26;
        end;
        if x > 0 then begin
            c := 64 + x;
            i := i + 1;
            t[i] := c;
        end;
        for x := 1 to i do
            ReturnText[x] := t[1 + i - x];
    end;

    [TryFunction]
    procedure GetIntegerRange(RangeName: Text; var FromRow: Integer; var FromCol: Integer; var ToRow: Integer; var ToCol: Integer);
    var
        FromText: Text;
        ToText: Text;
        ColumnCode: Code[3];
    begin
        FromRow := 0;
        FromCol := 0;
        ToRow := 0;
        ToCol := 0;

        RangeName := DelChr(RangeName, '=', '$');
        if Evaluate(FromRow, RangeName) then begin
            ToRow := FromRow;
            exit;
        end;

        if (StrPos(RangeName, ':') > 0) then begin
            FromText := CopyStr(RangeName, 1, StrPos(RangeName, ':') - 1);
            SplitCellAddress(FromText, ColumnCode, FromRow);
            if ColumnCode <> '' then
                FromCol := GetColumnInt(ColumnCode);
            RangeName := CopyStr(RangeName, StrPos(RangeName, ':') + 1);
        end;

        SplitCellAddress(RangeName, ColumnCode, ToRow);
        if ColumnCode <> '' then
            ToCol := GetColumnInt(ColumnCode);

        if (FromRow = 0) and (FromCol = 0) then begin
            FromRow := ToRow;
            FromCol := ToCol;
        end;
    end;

    procedure GetColumnInt(ColumnCode: Text[3]): Integer;
    var
        I: Integer;
        J: Integer;
        ReturnValue: Integer;
    begin
        ColumnCode := UpperCase(ColumnCode);

        if ColumnCode = '' then
            exit(0);

        J := 0;
        for I := StrLen(ColumnCode) downto 1 do begin
            if not (ColumnCode[I] in ['A' .. 'Z']) then
                Error(InvalidColumnCodeErr, ColumnCode);
            if J > 0 then
                ReturnValue += (ColumnCode[I] - 64) * Power(26, J)
            else
                ReturnValue := ReturnValue + (ColumnCode[I] - 64);
            J += 1;
        end;

        exit(ReturnValue);
    end;

    procedure SplitCellAddress(CellAddress: Text; var ColumnCode: Code[3]; var RowId: Integer);
    var
        I: Integer;
    begin
        ColumnCode := '';
        RowId := 0;

        if CellAddress = '' then
            exit;

        CellAddress := UpperCase(DelChr(CellAddress, '=', '$'));

        for I := 1 to StrLen(CellAddress) do begin
            if CellAddress[I] in ['1' .. '9'] then begin
                if not Evaluate(RowId, CopyStr(CellAddress, I)) then
                    Error(InvalidCellAddrErr, CellAddress);
                if I > 1 then
                    ColumnCode := CopyStr(CellAddress, 1, I - 1);
                exit;
            end;

            if not (CellAddress[I] in ['A' .. 'Z']) then
                Error(InvalidCellAddrErr, CellAddress);
        end;
        ColumnCode := CellAddress;
    end;

    procedure InitXMLNamespaceMngr(var XmlDoc: XmlDocument; var XmlNamespaceManager: XmlNamespaceManager)
    var
        RootElement: XmlElement;
        XmlAttr: XmlAttribute;
    begin
        Clear(XmlNamespaceManager);
        XmlDoc.GetRoot(RootElement);
        XmlNamespaceManager.NameTable(XmlDoc.NameTable());

        if RootElement.NamespaceUri <> '' then
            XmlNamespaceManager.AddNamespace('', RootElement.NamespaceUri);

        foreach XmlAttr IN RootElement.Attributes() DO
            if StrPos(XmlAttr.Name, 'xmlns:') = 1 then
                XmlNamespaceManager.AddNamespace(DelStr(XmlAttr.Name, 1, 6), XmlAttr.Value);
    end;

    procedure SelectSingleXmlElement(var XmlDoc: XmlDocument; XPath: Text; var FoundElement: XmlElement): Boolean
    var
        XmlNode: XmlNode;
        XmlNamespaceManager: XmlNamespaceManager;
    begin
        InitXMLNamespaceMngr(XmlDoc, XMLNamespaceManager);
        if not XMLDoc.SelectSingleNode(XPath, XmlNamespaceManager, XmlNode) then
            exit(false);
        FoundElement := XmlNode.AsXmlElement();
        exit(true);
    end;

    procedure AreXMLElementsTheSame(XmlElmnt: XmlElement; XmlElmnt2: XmlElement): Boolean
    begin
        exit(XmlElmnt.InnerXml.Replace('<x:', '<').Replace('xmlns:x', 'xmlns') = XmlElmnt2.InnerXml.Replace('<x:', '<').Replace('xmlns:x', 'xmlns'));
    end;

    procedure FindElementInCollection(SourceXmlElement: XmlElement; XmlElmnt: XmlElement; var ElmntId: integer): Boolean
    var
        XmlNode: XmlNode;
    begin
        ElmntId := 0;
        foreach XmlNode in SourceXmlElement.GetChildElements() do begin
            if AreXMLElementsTheSame(XmlNode.AsXmlElement(), XmlElmnt) then
                exit(true);
            ElmntId += 1;
        end;
        exit(false);
    end;

    procedure CreateElementWithAttribute(Name: Text; AttrName: Text; AttrValue: Text): XmlElement
    begin
        Exit(CreateElementWithAttribute(Name, MainNamespace(), AttrName, AttrValue));
    end;

    procedure CreateElementWithAttribute(Name: Text; Namespace: Text; AttrName: Text; AttrValue: Text): XmlElement
    var
        NewXmlElement: XmlElement;
    begin
        NewXmlElement := XmlElement.Create(Name, Namespace);
        NewXmlElement.SetAttribute(AttrName, AttrValue);
        Exit(NewXmlElement);
    end;



    procedure MainNamespace(): Text
    var
        MainNamespaceLbl: Label 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', Locked = true;
    begin
        exit(MainNamespaceLbl);
    end;

}
