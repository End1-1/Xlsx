#include "xlsxstyles.h"
#include <QString>

XlsxStyles::XlsxStyles()
{
    fZipFileName = "xl\\styles.xml";
    fStylesCount = 1;
    fFillsCount = 2;
    fFontsCount = 1;
    fHAlignsCount = 1;
}

void XlsxStyles::buildExcelData()
{
    fExcelData =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";

    fExcelData += QString("<fonts count=\"%1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font>").arg(fFontsCount);
    for (QMap<int, QFont>::const_iterator it = fFonts.begin(); it != fFonts.end(); it++) {
        fExcelData += "<font>";
        fExcelData += QString("<sz val=\"%1\"/>").arg(it.value().pointSize());
        fExcelData += QString("<name val=\"%1\"/>").arg(it.value().family());
        if (it.value().bold()) {
            fExcelData += "<b/>";
        }
        fExcelData += "</font>";
    }
    fExcelData += "</fonts>";

    fExcelData += QString("<fills count=\"%1\">").arg(fFillsCount);
    fExcelData += "<fill><patternFill patternType=\"none\"/></fill>";
    fExcelData += "<fill><patternFill patternType=\"gray125\"/></fill>";
    for (QMap<int, QColor>::const_iterator it = fBackgroundFills.begin(); it != fBackgroundFills.end(); it++) {
        fExcelData += QString("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"%1\"/><bgColor rgb=\"%1\"/></patternFill></fill>")
                .arg(it.value().name().replace("#", "ff"));
    }
    fExcelData += "</fills>";

    fExcelData += "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border>";
    fExcelData += "</borders>";

    fExcelData += "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>";
    fExcelData += "</cellStyleXfs>";

    fExcelData += QString("<cellXfs count=\"%1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>").arg(fStylesCount);
    for (QMap<int, QString>::const_iterator it = fStyles.begin(); it != fStyles.end(); it++) {
        int numFmtId = fNumFmtId[it.key()];
        int fontId = fFontsMap[it.value()];
        int borderId = 0;
        int fillId = fBackgroundFillsMap[it.value()];
        int xfId = 0;
        fExcelData += QString("<xf numFmtId=\"%1\" fontId=\"%2\" fillId=\"%3\" borderId=\"%4\" xfId=\"%5\">")
                .arg(numFmtId)
                .arg(fontId)
                .arg(fillId)
                .arg(borderId)
                .arg(xfId);
        if (fHAlign.contains(it.key())) {
            fExcelData += QString("<alignment horizontal=\"%1\"/>").arg(fHAlign[it.key()]);
        }
        fExcelData += "</xf>";
    }
    fExcelData += "</cellXfs>";

    fExcelData += "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>";
    fExcelData += "</styleSheet>";
}

int XlsxStyles::styleNum(const QString &name)
{
    return fStyles.key(name);
}

QString XlsxStyles::styleName(int style)
{
    return fStyles[style];
}

int XlsxStyles::copyStyle(const QString &srcName, const QString &newStyleName)
{
    if (fStyles.values().contains(newStyleName)) {
        return fStyles.key(newStyleName);
    }
    int styleCount = fStylesCount;
    if (fBackgroundFillsMap.contains(srcName)) {
        addBackgrounFill(newStyleName, fBackgroundFills[fBackgroundFillsMap[srcName]]);
    }
    if (fFontsMap.contains(srcName)) {
        addFont(newStyleName, fFonts[fFontsMap[srcName]]);
    }
    if (styleCount == fStylesCount) {
        fStyles[fStylesCount++] = newStyleName;
    }
    return styleNum(newStyleName);
}

void XlsxStyles::addFont(const QString &name, const QFont &font)
{
    checkStyleNumber(name);
    if (!fFontsMap.contains(name)) {
        fFontsMap[name] = fFontsCount++;
    }
    fFonts[fFontsMap[name]] = font;
}

void XlsxStyles::addBackgrounFill(const QString &name, const QColor &color)
{
    checkStyleNumber(name);
    if (!fBackgroundFillsMap.contains(name)) {
        fBackgroundFillsMap[name] = fFillsCount++;
    }
    fBackgroundFills[fBackgroundFillsMap[name]] = color;
}

void XlsxStyles::addNumFmtId(const QString &name, int id)
{
    fNumFmtId[styleNum(name)] = id;
}

void XlsxStyles::addHAlignment(const QString &name, const QString &aligment)
{
    checkStyleNumber(name);
    if (!fHAlignMap.contains(name)) {
        fHAlignMap[name] = fHAlignsCount++;
    }
    fHAlign[fHAlignMap[name]] = aligment;
}

void XlsxStyles::checkStyleNumber(const QString &name)
{
    if (!fStyles.values().contains(name)) {
        fStyles[fStylesCount++] = name;
    }
}
