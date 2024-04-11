```
var
    xlBuf: Record "Excel Buffer Extended" temporary;
    I: Integer;
    J: Integer;
begin

    xlBuf.Reset();
    xlBuf.DeleteAll();

    xlBuf.CreateBook('Sheet1');
    xlBuf.ClearNewRow();
    xlBuf.Validate("Row No.", 1);
    xlBuf.Validate("Column No.", 1);
    xlBuf."Cell Value as Text" := 'Hello, world!';
    xlBuf."Font Name" := 'Tahoma';
    xlBuf."Font Color" := 'FFFF0000'; //Red
    xlBuf."Font Size" := 14;
    xlBuf."Horizontal Alignment" := xlBuf."Horizontal Alignment"::Center;
    xlBuf."Right Border Style" := xlBuf."Right Border Style"::Thick;
    xlBuf."Right Border Color" := 'FF92D050';  // Green
    xlBuf."Background Color" := 'FFEEEEEE';
    xlBuf.Insert;
    xlBuf.WriteSheet('Test', 'Company Name', 'UserID');
    // WriteSheet adds Xml file to a zip-archive, all worksheet modifications must be made before calling this function

    xlBuf.DeleteAll();
    xlBuf.AddNewSheet('Sheet2');
    xlBuf.ClearNewRow();

    xlBuf.NewRow();
    xlBuf.AddColumn('Long text to check shrinking', false, '', false, false, false, '', xlBuf."Cell Type"::Text);
    xlBuf."Background Color" := 'FF92D050';
    xlBuf."Shrink To Fit" := true;
    xlBuf.Modify();

    xlBuf.AddColumn('', false, xlBuf."Cell Type"::Text, '');
    xlBuf."Patern Style" := xlBuf."Patern Style"::"Thick Diagonal Crosshatch";
    //xlBuf."Patern Color" := 'FF00B0F0';
    xlBuf."Background Color" := 'FFFF0000';
    xlBuf.Modify();

    xlBuf.AddColumn('', false, xlBuf."Cell Type"::Text, '');
    xlBuf."Gradient Color 1" := 'FFFF0000';
    xlBuf."Gradient Color 2" := 'FF92D050';
    xlBuf."Shading Style" := xlBuf."Shading Style"::Horizontal;
    xlBuf.Modify();

    xlBuf.AddColumn('', false, xlBuf."Cell Type"::Text, '');
    xlBuf."Gradient Color 1" := 'FFFF0000';
    xlBuf."Gradient Color 2" := 'FF92D050';
    xlBuf."Shading Style" := xlBuf."Shading Style"::"Vertical Middle";
    xlBuf.Modify();

    xlBuf.NewRow();
    xlBuf.AddColumn('', false, xlBuf."Cell Type"::Text, '');
    xlBuf."Border Style" := xlBuf."Border Style"::Thick;
    xlBuf.Modify();

    xlBuf.AddColumn('', false, xlBuf."Cell Type"::Text, '');
    xlBuf."Diagonal Border Style" := xlBuf."Border Style"::Thick;
    xlBuf."Diagonal Border Color" := 'FF92D050';
    xlBuf.Modify();

    xlBuf.AddColumn('', false, xlBuf."Cell Type"::Text, '');
    xlBuf."Diagonal Border Style" := xlBuf."Border Style"::Thick;
    xlBuf."Diagonal Border Type" := xlBuf."Diagonal Border Type"::"Up and Down";
    xlBuf.Modify();

    xlBuf.WriteSheet('', '', '');

    xlBuf.DeleteAll();
    xlBuf.AddNewSheet('Sheet3');
    xlBuf.ClearNewRow();

    xlBuf.AddColumn('First column width is 250px', false, xlBuf."Cell Type"::Text, '');
    xlBuf.SetColumnWidth(1, 250);

    xlBuf.AddColumn('Second column is hidden', false, xlBuf."Cell Type"::Text, '');
    xlBuf.SetColumnHidden(2, true);
    // or xlBuf.SetColumnsProperty('B',xlBuf.Property::Hidden,TRUE);

    // group columns
    xlBuf.SetColumnsOutlineLevel('E:F', 1);
    // or xlBuf.SetColumnsProperty('5:6',xlBuf.Property::"Outline Level",1);

    xlBuf.SetCurrent(3, 0);
    xlBuf.AddColumn('Third row height is 30px', false, xlBuf."Cell Type"::Text, 'Vertical Alignment:Center');
    xlBuf.SetRowHeight(3, 30);

    xlBuf.SetCurrent(4, 0);
    xlBuf.AddColumn('Cell with comment', false, xlBuf."Cell Type"::Text, 'Comment: test comment');
    xlBuf.AddColumn('Cell with comment2', false, xlBuf."Cell Type"::Text, 'Comment: another comment');

    // Group and hide rows
    xlBuf.SetRowsOutlineLevel(4, 6, 1);
    xlBuf.SetRowsHidden(4, 6, true);
    xlBuf.SetRowsSummaryAbove(true);
    // Add page header and footer
    xlBuf.SetPageHeaderFooterSettings(false, true);
    xlBuf.AddPageHeaderFooter("Excel Page HeaderFooter Type"::firstHeader, StrSubstNo('&L%1', UserId));  // will be displayed on the first page
    xlBuf.AddPageHeaderFooter("Excel Page HeaderFooter Type"::oddHeader, StrSubstNo('&R%1', 'Some text'));  // will be displayed on other pages
    xlBuf.AddPageHeaderFooter("Excel Page HeaderFooter Type"::oddFooter, 'Left', '&BCenter', 'Right');
    xlBuf.WriteSheet('', '', '');


    xlBuf.DeleteAll();
    xlBuf.AddNewSheet('Sheet4');
    xlBuf.ClearNewRow();

    xlBuf.SetDefaultProperties(
      'Bold;' +
      'Top Border Style:Medium;' +
      'Bottom Border Style:Medium;' +
      'Left Border Style:Thin;' +
      'Right Border Style:Thin;');

    for I := 1 to 5 do
        case I of
            1:
                xlBuf.AddColumn(StrSubstNo('Column Header %1', I), false, xlBuf."Cell Type"::Text, 'Left Border Style:Medium');
            5:
                xlBuf.AddColumn(StrSubstNo('Column Header %1', I), false, xlBuf."Cell Type"::Text, 'Right Border Style:Medium');
            else
                xlBuf.AddColumn(StrSubstNo('Column Header %1', I), false, xlBuf."Cell Type"::Text, '');
        end;
    xlBuf.AddAutoFilter(StrSubstNo('A%1:%2%1', xlBuf."Row No.", xlBuf.xlColID));
    xlBuf.SetDefaultProperties('Top Border Style: Thin; Bottom Border Style: Thin');

    for J := 2 to 4 do begin
        xlBuf.NewRow;
        for I := 1 to 5 do
            case I of
                1:
                    xlBuf.AddColumn(StrSubstNo('Cell R%1C%2', J, I), false, xlBuf."Cell Type"::Text, 'Left Border Style:Medium;Horizontal Alignment:Center');
                5:
                    xlBuf.AddColumn(StrSubstNo('Cell R%1C%2', J, I), false, xlBuf."Cell Type"::Text, 'Right Border Style:Medium;Horizontal Alignment:Right');
                else
                    xlBuf.AddColumn(StrSubstNo('Cell R%1C%2', J, I), false, xlBuf."Cell Type"::Text, '');
            end;
    end;

    // Footer is without borders, it has only bold text. Top border is necessary to finish the border of the table body
    xlBuf.SetDefaultProperties('Bold;Top Border Style:Medium');
    xlBuf.NewRow;
    xlBuf.AddColumn('Rows count:', false, xlBuf."Cell Type"::Text, '');
    xlBuf.AddColumn('', false, xlBuf."Cell Type"::Text, '');
    xlBuf.AddColumn('', false, xlBuf."Cell Type"::Text, '');
    xlBuf.AddColumn('', false, xlBuf."Cell Type"::Text, '');
    xlBuf.MergeCells(StrSubstNo('A%1:D%1', xlBuf."Row No."));

    xlBuf.SetCurrent(xlBuf."Row No.", 4);
    xlBuf.AddColumn(xlBuf."Row No." - 1, false, xlBuf."Cell Type"::Number, '');
    xlBuf.FreezeTopRow();
    xlBuf.WriteSheet('', '', UserId);


    xlBuf.CloseBook();
    xlBuf.SetFriendlyFilename('NewTest');
    xlBuf.DownloadExcelFile();
end;
```
