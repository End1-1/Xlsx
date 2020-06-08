#ifndef XLSXSHEET_H
#define XLSXSHEET_H

#include "xlsx.h"
#include <QMap>
#include <QVariant>

class XlsxCell;
class XlsxSharedString;
class XlsxStyles;

class XlsxSheet : public Xlsx
{
public:
    XlsxSheet(const QString &name);
    virtual void buildExcelData();
    XlsxSharedString *fSharedStrings;
    XlsxStyles *fStyle;
    QString &name();
    XlsxCell *addCell(int row, int column, QVariant cellValue, int style = 0);
    void setColumnWidth(int column, int width);
    void setSpan(const QString &span);
    void setSpan(const QString &f, const QString &s, int row);

private:
    QMap<int, QMap<int, XlsxCell *> > fCells;
    QMap<int, int> fColumnWidths;
    QStringList fSpan;
    QString fName;
};

#endif // XLSXSHEET_H
