#ifndef XLSXCELL_H
#define XLSXCELL_H

#include <QVariant>
#include <QString>

class XlsxSharedString;

class XlsxCell
{
public:
    XlsxCell(int row, int column, const QVariant &cellValue);
    QVariant &value();
    const QString &address();
    QString fValueType;
    int fStyle;

private:
    int fRow;
    int fColumn;
    QString fAddress;
    QVariant fCellValue;
};

#endif // XLSXCELL_H
