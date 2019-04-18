#include "xlsxsheet.h"
#include "xlsxcell.h"
#include "xlsxsharedstring.h"
#include <QString>

XlsxSheet::XlsxSheet(const QString &name) :
    Xlsx()
{
    fName = name;
    fZipFileName = "xl/worksheets/" + name + ".xml";
}

void XlsxSheet::buildExcelData()
{
    fExcelData =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";
    fExcelData += "<cols>";
    for (QMap<int, int>::const_iterator it = fColumnWidths.begin(); it != fColumnWidths.end(); it++) {
        fExcelData += QString("<col min=\"%1\" max=\"%1\" width=\"%2\" costumWidth=\"1\"/>")
                .arg(it.key())
                .arg(it.value());
    }
    fExcelData += "</cols><sheetData>";
    for (QMap<int, QMap<int, XlsxCell *> >::const_iterator ir = fCells.begin(); ir != fCells.end(); ir++) {
        fExcelData += QString("<row r=\"%1\">").arg(ir.key());
        for (QMap<int, XlsxCell *>::const_iterator ic = ir.value().begin(); ic != ir.value().end(); ic++) {
            fExcelData += QString("<c s=\"%1\" r=\"%2\" %3><v>%4</v></c>")
                    .arg(ic.value()->fStyle)
                    .arg(ic.value()->address())
                    .arg(ic.value()->fValueType)
                    .arg(ic.value()->value().toString());
        }
        fExcelData += "</row>";
    }
    fExcelData += "</sheetData></worksheet>";
}

QString &XlsxSheet::name()
{
    return fName;
}

XlsxCell *XlsxSheet::addCell(int row, int column, QVariant cellValue)
{
    QString vt;
    switch (cellValue.type()) {
    case QVariant::String:
        cellValue = fSharedStrings->addString(cellValue.toString());
        vt = " t=\"s\"";
        break;
    default:
        break;
    }
    XlsxCell *c = new XlsxCell(row, column, cellValue);
    c->fValueType = vt;
    fCells[row][column] = c;
    return c;
}

void XlsxSheet::setColumnWidth(int column, int width)
{
    fColumnWidths[column] = width;
}
