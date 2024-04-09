namespace ASD.Excel;

enum 50101 "Excel Page Orientation"
{
    Extensible = true;

    value(0; Default)
    {
        Caption = 'Default';
    }
    value(1; Portrait)
    {
        Caption = 'Portrait';
    }
    value(2; Landscape)
    {
        Caption = 'Landscape';
    }
}
