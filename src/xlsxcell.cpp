#include "xlsxcell.h"

static const char alphabet[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

XlsxCell::XlsxCell(int row, int column, const QVariant &cellValue)
{
    fRow = row;
    fColumn = column;
    fCellValue = cellValue;
    fAddress = calculateAddress(row, column);
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

QString XlsxCell::calculateAddress(int row, int col)
{
    QString a;
    while (col > 0) {
        col--;  // Для того чтобы правильно работать с системой счисления от 1
        a.prepend(alphabet[col % 26]);  // Добавляем букву в начало строки
        col /= 26;  // Переходим к следующему разряду
    }
    return a + QString::number(row);
}

int XlsxCell::style()
{
    return fStyle;
}

void XlsxCell::setStyle(int style)
{
    fStyle = style;
}
