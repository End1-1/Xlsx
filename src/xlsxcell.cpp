#include "xlsxcell.h"

static const char alphabet[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

XlsxCell::XlsxCell(int row, int column, const QVariant &cellValue)
{
    fRow = row;
    fColumn = column;
    do  {
        int m = (column - 1) % 25;
        fAddress += alphabet[m];
        column -= 26;
    } while (column > 26);
    fCellValue = cellValue;
    fAddress += QString::number(row);
    fStyle = 0;
}

QVariant &XlsxCell::value()
{
    return fCellValue;
}

const QString &XlsxCell::address()
{
    return fAddress;
}

int XlsxCell::style()
{
    return fStyle;
}

void XlsxCell::setStyle(int style)
{
    fStyle = style;
}
