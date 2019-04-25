#include "xlsxcell.h"

static const char alphabet[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

XlsxCell::XlsxCell(int row, int column, const QVariant &cellValue)
{
    fRow = row;
    fColumn = column;
    do  {
        int m = (column) / 27;
        if (m == 0) {
            m = column % 27;
        }
        column -= 26;
        fAddress += alphabet[m - 1];
    } while (column > 0);
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
