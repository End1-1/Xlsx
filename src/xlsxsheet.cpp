#include "xlsxsheet.h"
#include "xlsxcell.h"
#include "xlsxsharedstring.h"
#include "xlsxstyles.h"
#include <QString>
#include <QDate>

static const int NUMFMID_DATE = 14;
static const int NUMFMID_DATETIME = 22;

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
        fExcelData += QString("<col min=\"%1\" max=\"%1\" width=\"%2\" %3costumWidth=\"1\"/>")
                .arg(it.key())
                .arg(it.value())
                .arg(it.value() == 0 ? "hidden=\"1\" " : "");
    }
    fExcelData += "</cols><sheetData>";
    for (QMap<int, QMap<int, XlsxCell *> >::const_iterator ir = fCells.begin(); ir != fCells.end(); ir++) {
        fExcelData += QString("<row r=\"%1\">").arg(ir.key());
        for (QMap<int, XlsxCell *>::const_iterator ic = ir.value().begin(); ic != ir.value().end(); ic++) {
            fExcelData += QString("<c s=\"%1\" r=\"%2\"%3><v>%4</v></c>")
                    .arg(ic.value()->style())
                    .arg(ic.value()->address())
                    .arg(ic.value()->fValueType.isEmpty() ? "" : ic.value()->fValueType)
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

XlsxCell *XlsxSheet::addCell(int row, int column, QVariant cellValue, int style)
{
    QString vt;
    switch (cellValue.type()) {
    case QVariant::String:
        cellValue = fSharedStrings->addString(cellValue.toString()
                                              .replace("#", "===|||+++patternchik")
                                              .replace("&", "&#38;")
                                              .replace("===|||+++patternchik", "&#35;")
                                              .replace("*", "&#42;")
                                              .replace("/", "&#47;")
                                              .replace("<", "&lt;")
                                              .replace(">", "&gt;")
                                              .replace("\"", "&quot;")
                                              .replace("'", "&apos;"));
        vt = " t=\"s\"";
        break;
    case QVariant::Date: {
        cellValue = QDate::fromString("1900-01-01", "yyyy-MM-dd").daysTo(cellValue.toDate()) + 2;
        QString newStyle = fStyle->styleName(style) + "withdate";
        int newStyleNum = fStyle->copyStyle(fStyle->styleName(style), newStyle);
        fStyle->addNumFmtId(newStyle, NUMFMID_DATE);
        style = newStyleNum;
        break;
    }
    case QVariant::DateTime: {
        /*
        cellValue = QDate::fromString("1900-01-01", "yyyy-MM-dd").daysTo(cellValue.toDate()) + 1 + (QTime::fromString("00:00:00").secsTo(cellValue.toDateTime().time()) * 0.00001574074074074073);
        QString newStyle = fStyle->styleName(style) + "withdatetime";
        int newStyleNum = fStyle->copyStyle(fStyle->styleName(style), newStyle);
        fStyle->addNumFmtId(newStyle, NUMFMID_DATETIME);
        style = newStyleNum;
        break;
        */
        cellValue = fSharedStrings->addString(cellValue.toDateTime().toString("dd/MM/yyyy HH:mm:ss").replace("/", "&#47;"));
        vt = " t=\"s\"";
        break;
    }
    default:
        break;
    }
    XlsxCell *c = new XlsxCell(row, column, cellValue);
    c->fValueType = vt;
    if (style > 0) {
        c->setStyle(style);
    }
    fCells[row][column] = c;
    return c;
}

void XlsxSheet::setColumnWidth(int column, int width)
{
    fColumnWidths[column] = width;
}
