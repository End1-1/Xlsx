#ifndef XLSXSHEET_H
#define XLSXSHEET_H

#include "xlsx.h"
#include <QMap>
#include <QVariant>

class XlsxCell;
class XlsxSharedString;

class XlsxSheet : public Xlsx
{
public:
    XlsxSheet(const QString &name);
    virtual void buildExcelData();
    XlsxSharedString *fSharedStrings;
    QString &name();
    XlsxCell *addCell(int row, int column, QVariant cellValue);
    void setColumnWidth(int column, int width);

private:
    QMap<int, QMap<int, XlsxCell *> > fCells;
    QMap<int, int> fColumnWidths;
    QString fName;
};

#endif // XLSXSHEET_H
