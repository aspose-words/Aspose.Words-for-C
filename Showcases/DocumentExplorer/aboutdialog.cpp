#include "aboutdialog.h"
#include "ui_aboutdialog.h"

#include "hidpisupport.h"

AboutDialog::AboutDialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::AboutDialog)
{
    ui->setupUi(this);

    ui->labelLogo->setMaximumSize(ui->labelLogo->maximumSize() * HiDpiSupport::getImageScaleFactor());
}

AboutDialog::~AboutDialog()
{
    delete ui;
}
