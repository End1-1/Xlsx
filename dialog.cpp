#include "dialog.h"
#include "ui_dialog.h"
#include "xlsxwriter.h"
#include "xlsxdocument.h"
#include "xlsxworkbook.h"
#include "xlsxstyles.h"
#include "xlsxcell.h"
#include "xlsxsheet.h"
#include <QMessageBox>

Dialog::Dialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::Dialog)
{
    ui->setupUi(this);
}

Dialog::~Dialog()
{
    delete ui;
}

void Dialog::on_pushButton_clicked()
{
    QString err;
    XlsxDocument d;
    XlsxWorkBook *wb = d.addWorkbook();
    XlsxSheet *sh = wb->addSheet("Sheet1");
    XlsxStyles *s = d.style();
    QFont f(qApp->font());
    f.setFamily("Arial");
    f.setPointSize(9);
    f.setBold(true);
    s->addFont("data", f);
    s->addBackgrounFill("header", QColor::fromRgb(255, 255, 0));
    sh->addCell(1, 1, 1)->fStyle = d.style()->styleNum("header");
    sh->addCell(1, 2, 9.01)->fStyle = d.style()->styleNum("data");
    sh->addCell(4, 2, 555);
    sh->addCell(4, 23, 26)->fStyle = d.style()->styleNum("data");
    sh->addCell(5, 24, 27);
    sh->addCell(6, 1, "TEST");
    sh->addCell(6, 2, "TEST2");
    sh->addCell(6, 3, "TESTssss");
    sh->setColumnWidth(1, 30);
    sh->setColumnWidth(2, 60);
    if (!d.save("d:/temp/test.xlsx", err)) {
        QMessageBox::critical(this, "Error", err);
    }
}

void Dialog::on_pushButton_2_clicked()
{
    QString err;
    XlsxWriter xw("d:/temp/test33file.zip");
    char d1[] = "data1";
    char d2[] = "second data";
    char d3[] = "data n3";
    xw.writeLocalHeader("1.txt", reinterpret_cast<const uint8_t *>(d1), strlen(d1));
    xw.writeLocalHeader("2.txt", reinterpret_cast<const uint8_t *>(d2), strlen(d2));
    xw.writeLocalHeader("3.txt", reinterpret_cast<const uint8_t *>(d3), strlen(d3));
    xw.writeEOD();
}
