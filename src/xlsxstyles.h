#ifndef XLSXSTYLES_H
#define XLSXSTYLES_H

#include "xlsx.h"
#include <QColor>
#include <QFont>

class XlsxStyles : public Xlsx
{
public:
    XlsxStyles();
    virtual void buildExcelData();
    int styleNum(const QString &name);
    void addFont(const QString &name, const QFont &font);
    void addBackgrounFill(const QString &name, const QColor &color);

private:
    QMap<int, QString> fStyles;
    QMap<QString, int> fBackgroundFillsMap;
    QMap<int, QColor> fBackgroundFills;
    QMap<QString, int> fFontsMap;
    QMap<int, QFont> fFonts;
    int fStylesCount;
    int fFontsCount;
    int fFillsCount;
    void checkStyleNumber(const QString &name);
};

#endif // XLSXSTYLES_H
