#ifndef XLSXSTYLES_H
#define XLSXSTYLES_H

#include "xlsx.h"
#include <QColor>
#include <QFont>

#define xls_alignment_center "center"

class XlsxStyles : public Xlsx
{
public:
    XlsxStyles();
    virtual void buildExcelData();
    int styleNum(const QString &name);
    QString styleName(int style);
    int copyStyle(const QString &srcName, const QString &newStyleName);
    void addFont(const QString &name, const QFont &font);
    void addBackgrounFill(const QString &name, const QColor &color);
    void addNumFmtId(const QString &name, int id);
    void addHAlignment(const QString &name, const QString &aligment);

private:
    QMap<int, QString> fStyles;
    QMap<QString, int> fBackgroundFillsMap;
    QMap<int, QColor> fBackgroundFills;
    QMap<QString, int> fFontsMap;
    QMap<int, QFont> fFonts;
    QMap<int, int> fNumFmtId;
    QMap<QString, int> fHAlignMap;
    QMap<int, QString> fHAlign;
    int fStylesCount;
    int fFontsCount;
    int fFillsCount;
    int fHAlignsCount;
    void checkStyleNumber(const QString &name);
};

#endif // XLSXSTYLES_H
