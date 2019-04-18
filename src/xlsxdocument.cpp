#include "xlsxdocument.h"
#include "xlsxwriter.h"
#include "xlsxworkbook.h"
#include "xlsxstyles.h"
#include "xlsxtheme.h"
#include "xlsxdocpropsapp.h"
#include "xlsxdocpropscore.h"
#include "xlsxsharedstring.h"
#include "xlsxcontenttype.h"
#include "xlsxrels.h"
#include "xlsxsheet.h"

XlsxDocument::XlsxDocument()
{
    fContentTypes = new XlsxContentType(this);
    fDocPropsApp = new XlsxDocPropsApp();
    fDocPropsCore = new XlsxDocPropsCore();
    fRels = new XlsxRels();
    fWorkBook = new XlsxWorkBook();
    fStyles = new XlsxStyles();
    fTheme = new XlsxTheme();
}

XlsxDocument::~XlsxDocument()
{
    if (fWorkBook) {
        delete fWorkBook;
    }
    delete fContentTypes;
    delete fDocPropsApp;
    delete fDocPropsCore;
    delete fStyles;
    delete fTheme;
}

XlsxWorkBook *XlsxDocument::addWorkbook()
{
    if (!fWorkBook) {
        fWorkBook = new XlsxWorkBook();
    }
    return fWorkBook;
}

XlsxWorkBook *XlsxDocument::workbook()
{
    return fWorkBook;
}

XlsxStyles *XlsxDocument::style()
{
    return fStyles;
}

bool XlsxDocument::save(const QString &fileName, QString &err)
{
    if (!fWorkBook) {
        err = "Workbook missing";
        return false;
    }
    XlsxWriter xw(fileName);
    if (!xw.fError.isEmpty()) {
        err = xw.fError;
        return false;
    }

    fContentTypes->writeToFile(&xw);
    fDocPropsCore->writeToFile(&xw);
    fDocPropsApp->writeToFile(&xw);
    fRels->writeToFile(&xw);
    fStyles->writeToFile(&xw);
    //fTheme->writeToFile(&xw);
    fWorkBook->writeToFile(&xw);
    for (int i = 0; i < fWorkBook->sheetsCount(); i++) {
        fWorkBook->sheet(i)->writeToFile(&xw);
    }
    xw.writeEOD();
    return true;
}
