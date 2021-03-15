#include "viewerwindow.h"
#include "ui_viewerwindow.h"

#include "hidpisupport.h"
#include "qtcorehelpers.h"

#include <Aspose.Words.Cpp/Rendering/PageInfo.h>
#include <drawing/bitmap.h>
#include <drawing/graphics.h>
#include <drawing/size.h>
#include <system/io/memory_stream.h>

#include <QFileDialog>
#include <QFileInfo>
#include <QInputDialog>
#include <QMessageBox>

using Aspose::QtHelpers::Convert;

namespace {

void scaleActionIcon(QAction* action)
{
    QPixmap pixmap = action->icon().pixmap(16, 16);
    action->setIcon(HiDpiSupport::scaleImage(pixmap));
}

} // unnamed namespace

ViewerWindow::ViewerWindow(System::SharedPtr<Aspose::Words::Document> document, QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::ViewerWindow),
    mDocument(document)
{
    ui->setupUi(this);

    scaleActionIcon(ui->actionOpen);
    scaleActionIcon(ui->actionFirstPage);
    scaleActionIcon(ui->actionPreviousPage);
    scaleActionIcon(ui->actionNextPage);
    scaleActionIcon(ui->actionLastPage);
    scaleActionIcon(ui->actionGoToPage);
    ui->toolBar->setIconSize(ui->toolBar->iconSize() * HiDpiSupport::getImageScaleFactor());

    mDocument->UpdatePageLayout();

    mPageNumber = 1;
    updatePage();
}

ViewerWindow::~ViewerWindow()
{
    delete ui;
}

void ViewerWindow::on_actionOpen_triggered()
{
    openDocument();
}

void ViewerWindow::on_actionExit_triggered()
{
    close();
}

void ViewerWindow::on_actionPreviousPage_triggered()
{
    mPageNumber--;
    updatePage();
}

void ViewerWindow::on_actionNextPage_triggered()
{
    mPageNumber++;
    updatePage();
}

void ViewerWindow::on_actionFirstPage_triggered()
{
    mPageNumber = 1;
    updatePage();
}

void ViewerWindow::on_actionLastPage_triggered()
{
    mPageNumber = mDocument->get_PageCount();
    updatePage();
}

void ViewerWindow::on_actionGoToPage_triggered()
{
    const int pageNumber = QInputDialog::getInt(this, "Go to Page", "Page number",
                                                mPageNumber, 1, mDocument->get_PageCount());
    if (pageNumber != mPageNumber)
    {
        mPageNumber = pageNumber;
        updatePage();
    }
}

// Shows the open file dialog box and opens a document.
void ViewerWindow::openDocument()
{
    const QString fileName = selectOpenFileName();
    if (fileName.isNull())
        return;

    // This operation can take some time so we set the Cursor to WaitCursor.
    setCursor(Qt::WaitCursor);
    QApplication::processEvents();

    // Load document is put in a try-catch block to handle situations when it fails for some reason.
    try
    {
        // Loads the document into Aspose.Words object model.
        mDocument = System::MakeObject<Aspose::Words::Document>(Convert(fileName));

        // Update UI
        setWindowTitle("Aspose.Words Rendering Demo - " + fileName);

        // Update page
        mPageNumber = 1;
        updatePage();
    }
    catch (const System::Exception& exception)
    {
        const QString exceptionMessage = Convert(exception.ToString());
        showErrorMessage(QString("Unable to load file %0. %1").arg(fileName).arg(exceptionMessage));
    }

    unsetCursor();
}

// Show selected page.
void ViewerWindow::updatePage()
{
    // This operation can take some time (for the first page) so we set the Cursor to WaitCursor.
    setCursor(Qt::WaitCursor);
    QApplication::processEvents();

    const bool canMoveBack = (mPageNumber > 1);
    ui->actionFirstPage->setEnabled(canMoveBack);
    ui->actionPreviousPage->setEnabled(canMoveBack);

    const bool canMoveForward = (mPageNumber < mDocument->get_PageCount());
    ui->actionLastPage->setEnabled(canMoveForward);
    ui->actionNextPage->setEnabled(canMoveForward);

    const int pageIndex = mPageNumber - 1;

    const auto pageInfo = mDocument->GetPageInfo(pageIndex);
    const float scale = HiDpiSupport::getImageScaleFactor();
    const auto imgSize = pageInfo->GetSizeInPixels(scale, 96.0f);

    const auto img = System::MakeObject<System::Drawing::Bitmap>(imgSize.get_Width(), imgSize.get_Height());
    {
        // Render page to System::Drawing::Bitmap.
        const auto gfx = System::Drawing::Graphics::FromImage(img);
        gfx->Clear(System::Drawing::Color::get_White());

        mDocument->RenderToScale(pageIndex, gfx, 0, 0, scale);
    }

    QPixmap pixmap;
    {
        // Copy data from System::Drawing::Bitmap to QPixmap.
        const auto stream = System::MakeObject<System::IO::MemoryStream>();
        img->Save(stream, System::Drawing::Imaging::ImageFormat::get_Bmp());

        stream->set_Position(0);

        const auto buffer = stream->GetBuffer();
        pixmap.loadFromData(buffer->data_ptr(), buffer->get_Length());
    }
    ui->imageLabel->setPixmap(pixmap);

    ui->imageLabel->adjustSize();
    ui->scrollAreaWidgetContents->adjustSize();

    ui->statusbar->showMessage(QString("Page %0 of %1").arg(mPageNumber).arg(mDocument->get_PageCount()));

    unsetCursor();
}

// Selects file name for a document to open or null.
QString ViewerWindow::selectOpenFileName()
{
    const QString filter =
        "All Supported Formats (*.doc *.dot *.docx *.dotx *.docm *.dotm *.xml *.wml *.rtf *.odt *.ott *.htm *.html *.xhtml *.mht *.mhtm *.mhtml);;"
        "Word 97-2003 Documents (*.doc *.dot);;"
        "Word 2007 OOXML Documents (*.docx *.dotx *.docm *.dotm);;"
        "XML Documents (*.xml *.wml);;"
        "Rich Text Format (*.rtf);;"
        "OpenDocument Text (*.odt *.ott);;"
        "Web Pages (*.htm *.html *.xhtml *.mht *.mhtm *.mhtml);;"
        "All Files (*.*)";
    const QString fileName = QFileDialog::getOpenFileName(this, "Open Document", mInitialDirectory, filter);
    if (!fileName.isNull())
    {
        const QFileInfo fileInfo(fileName);
        mInitialDirectory = fileInfo.dir().path();
    }
    return fileName;
}

// Show information about error.
void ViewerWindow::showErrorMessage(QString message)
{
    QMessageBox messageBox(QMessageBox::Critical,
                           "Document Explorer - unexpected error occured",
                           message, QMessageBox::Ok, this);
    messageBox.setTextInteractionFlags(Qt::TextSelectableByMouse);
    messageBox.exec();
}
