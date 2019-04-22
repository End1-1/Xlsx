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
    QString styleName(int style);
    int copyStyle(const QString &srcName, const QString &newStyleName);
    void addFont(const QString &name, const QFont &font);
    void addBackgrounFill(const QString &name, const QColor &color);
    void addNumFmtId(const QString &name, int id);

private:
    QMap<int, QString> fStyles;
    QMap<QString, int> fBackgroundFillsMap;
    QMap<int, QColor> fBackgroundFills;
    QMap<QString, int> fFontsMap;
    QMap<int, QFont> fFonts;
    QMap<int, int> fNumFmtId;
    int fStylesCount;
    int fFontsCount;
    int fFillsCount;
    void checkStyleNumber(const QString &name);
};

#endif // XLSXSTYLES_H
