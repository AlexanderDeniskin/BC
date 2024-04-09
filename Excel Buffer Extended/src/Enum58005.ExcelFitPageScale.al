namespace ASD.Excel;
enum 58005 "Excel Fit Page Scale"
{
    Extensible = true;

    value(0; "No scale")
    {
        Caption = 'No Scale';
    }
    value(1; Sheet)
    {
        Caption = 'Fit Sheet on One Page';
    }
    value(2; Columns)
    {
        Caption = 'Fit All Columns on One Page';
    }
    value(3; Rows)
    {
        Caption = 'Fot All Rows on One Page';
    }
}