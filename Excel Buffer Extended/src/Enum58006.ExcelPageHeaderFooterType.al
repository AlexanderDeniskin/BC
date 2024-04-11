namespace ASD.Excel;

enum 58006 "Excel Page HeaderFooter Type"
{
    Extensible = true;

    value(0; oddHeader)
    {
        Caption = 'oddHeader';
    }
    value(1; oddFooter)
    {
        Caption = 'oddFooter';
    }
    value(2; evenHeader)
    {
        Caption = 'evenHeader';
    }
    value(3; evenFooter)
    {
        Caption = 'evenFooter';
    }
    value(4; firstHeader)
    {
        Caption = 'firstHeader';
    }
    value(5; firstFooter)
    {
        Caption = 'firstFooter';
    }
}
