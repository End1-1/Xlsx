#include "xlsxdocpropscore.h"

XlsxDocPropsCore::XlsxDocPropsCore() :
    Xlsx()
{
    fZipFileName = "docProps\\core.xml";
    fExcelData =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<cp:coreProperties "
            "xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
            "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" "
            "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" "
            "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
            "<dc:creator>Administrator</dc:creator>"
            "<cp:lastModifiedBy>Administrator</cp:lastModifiedBy>"
            "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2019-04-15T09:43:27Z</dcterms:created>"
            "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">2019-04-15T09:43:43Z</dcterms:modified>"
            "</cp:coreProperties>";
}
