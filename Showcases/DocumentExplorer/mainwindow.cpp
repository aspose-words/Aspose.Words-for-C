#include "mainwindow.h"
#include "./ui_mainwindow.h"

#include "aboutdialog.h"
#include "viewerwindow.h"

#include "hidpisupport.h"
#include "qtcorehelpers.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOptions.h>
#include <Aspose.Words.Cpp/Licensing/License.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>

#include <QDir>
#include <QFileDialog>
#include <QMessageBox>
#include <QShortcut>
#include <QScreen>

#include <map>

using Aspose::QtHelpers::Convert;

namespace {

void scaleActionIcon(QAction* action)
{
    QPixmap pixmap = action->icon().pixmap(16, 16);
    action->setIcon(HiDpiSupport::scaleImage(pixmap));
}

} // unnamed namespace

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    scaleActionIcon(ui->actionOpen);
    scaleActionIcon(ui->actionSaveAs);
    scaleActionIcon(ui->actionRender);
    scaleActionIcon(ui->actionRemoveNode);
    scaleActionIcon(ui->actionExpandAll);
    scaleActionIcon(ui->actionCollapseAll);
    ui->toolBar->setIconSize(ui->toolBar->iconSize() * HiDpiSupport::getImageScaleFactor());

    findAndApplyLicense();

    QShortcut* shortcut = new QShortcut(QKeySequence(Qt::Key_Delete), ui->treeView);
    connect(shortcut, &QShortcut::activated, this, &MainWindow::on_actionRemoveNode_triggered);

    resize(QGuiApplication::primaryScreen()->availableSize() * 3 / 5);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_actionOpen_triggered()
{
    openDocument();
}

void MainWindow::on_actionSaveAs_triggered()
{
    saveDocument();
}

void MainWindow::on_actionRender_triggered()
{
    renderDocument();
}

void MainWindow::on_actionExit_triggered()
{
    close();
}

void MainWindow::on_actionRemoveNode_triggered()
{
    removeNode();
}

void MainWindow::on_actionExpandAll_triggered()
{
    expandAllNodes();
}

void MainWindow::on_actionCollapseAll_triggered()
{
    collapseAllNodes();
}

void MainWindow::on_actionAbout_triggered()
{
    AboutDialog aboutDialog(this);
    aboutDialog.setWindowFlags(Qt::Dialog | Qt::CustomizeWindowHint | Qt::WindowTitleHint);
    aboutDialog.window()->layout()->setSizeConstraint(QLayout::SetFixedSize);
    aboutDialog.exec();
}

void MainWindow::treeView_currentChanged(const QModelIndex &current, const QModelIndex &/*previous*/)
{
    // Enable/disable 'Remove Node' action.
    ui->actionRemoveNode->setEnabled(mDocumentModel->canRemove(current));

    // Enable/disable 'Expand All' and 'Collapse All' actions.
    const bool hasChildren = mDocumentModel->hasChildren(current);
    ui->actionExpandAll->setEnabled(hasChildren);
    ui->actionCollapseAll->setEnabled(hasChildren);

    // Show the text contained by selected document node.
    ui->textEdit->setHtml(getNodeText(current));
}

// Search for Aspose.Words license in the application directory.
void MainWindow::findAndApplyLicense()
{
    // Try to find Aspose.Total.Cpp license.
    QString licenseFile = QApplication::applicationDirPath() + QDir::separator() + "Aspose.Total.Cpp.lic";

    if (QFile::exists(licenseFile))
    {
        licenseAsposeWords(licenseFile);
        return;
    }

    licenseFile = QApplication::applicationDirPath() + QDir::separator() + "Aspose.Words.Cpp.lic";
    if (QFile::exists(licenseFile))
        licenseAsposeWords(licenseFile);
}

// This code activates Aspose.Words license.
// If you don't specify a license, Aspose.Words will work in evaluation mode.
void MainWindow::licenseAsposeWords(QString licenseFile)
{
    auto licenseWords = System::MakeObject<Aspose::Words::License>();
    licenseWords->SetLicense(Convert(licenseFile));
}

// Shows the open file dialog box and opens a document.
void MainWindow::openDocument()
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

        // Detach old model
        ui->treeView->setModel(nullptr);

        // Create new model (and delete old)
        mDocumentModel = std::make_unique<DocumentModel>(mDocument);

        // Attach new model
        ui->treeView->setModel(mDocumentModel.get());

        // Update UI
        setWindowTitle("Document Explorer - " + fileName);

        ui->actionSaveAs->setEnabled(true);
        ui->actionRender->setEnabled(true);
        ui->actionExpandAll->setEnabled(true);
        ui->actionCollapseAll->setEnabled(true);
        ui->actionRemoveNode->setEnabled(true);

        connect(ui->treeView->selectionModel(), &QItemSelectionModel::currentChanged, this, &MainWindow::treeView_currentChanged);

        // Shows the immediate children of the document node.
        ui->treeView->expandToDepth(0);

        // Selects root node of the tree.
        ui->treeView->setCurrentIndex(mDocumentModel->index(0, 0));
    }
    catch (const System::Exception& exception)
    {
        showErrorMessage(Convert(exception.ToString()));
    }

    unsetCursor();
}

// Saves the document with the name and format provided in standard Save As dialog.
void MainWindow::saveDocument()
{
    if (mDocument == nullptr)
        return;

    const QString fileName = selectSaveFileName();
    if (fileName.isNull())
        return;

    // This operation can take some time so we set the Cursor to WaitCursor.
    setCursor(Qt::WaitCursor);
    QApplication::processEvents();

    // This operation is put in try-catch block to handle situations when operation fails for some reason.
    try
    {
        auto saveOptions = Aspose::Words::Saving::SaveOptions::CreateSaveOptions(Convert(fileName));
        saveOptions->set_PrettyFormat(true);

        // For most of the save formats it is enough to just invoke Aspose.Words save.
        mDocument->Save(Convert(fileName), saveOptions);
    }
    catch (const System::Exception& exception)
    {
        showErrorMessage(Convert(exception.ToString()));
    }

    unsetCursor();
}

void MainWindow::renderDocument()
{
    if (mDocument == nullptr)
        return;

    auto viewerWindow = new ViewerWindow(mDocument, this);
    viewerWindow->setAttribute(Qt::WA_DeleteOnClose);
    viewerWindow->setWindowModality(Qt::WindowModal);
    viewerWindow->showMaximized();
}

// Removes currently selected node.
void MainWindow::removeNode()
{
    auto index = ui->treeView->currentIndex();
    if (!index.isValid())
        return;

    mDocumentModel->removeRows(index.row(), 1, index.parent());

    // Update GUI after removing node
    index = ui->treeView->currentIndex();
    treeView_currentChanged(index, index);
}

// Expands all nodes under currently selected node.
void MainWindow::expandAllNodes()
{
    // This operation can take some time so we set the Cursor to WaitCursor.
    setCursor(Qt::WaitCursor);
    QApplication::processEvents();

    auto index = ui->treeView->currentIndex();
    if (index.isValid())
        ui->treeView->expandRecursively(index);

    unsetCursor();
}

// Collapses all nodes under currently selected node.
void MainWindow::collapseAllNodes()
{
    // This operation can take some time so we set the Cursor to WaitCursor.
    setCursor(Qt::WaitCursor);
    QApplication::processEvents();

    auto index = ui->treeView->currentIndex();
    if (index.isValid())
        collapseNodesRecursively(index);

    unsetCursor();
}

void MainWindow::collapseNodesRecursively(const QModelIndex& parent)
{
    int count = mDocumentModel->rowCount(parent);
    if (count > 0)
    {
        for (int i = 0; i < count; ++i)
            collapseNodesRecursively(mDocumentModel->index(i, 0, parent));

        ui->treeView->collapse(parent);
    }
}

// Get text contained by the corresponding document node.
QString MainWindow::getNodeText(const QModelIndex& index) const
{
    using Aspose::Words::ControlChar;

    // Map of character to string that we use to display MS Word control characters.
    static const std::map<char16_t, QString> controlCharacters = {
        { ControlChar::CellChar, "[!Cell!]" },
        { ControlChar::ColumnBreakChar, "[!ColumnBreak!]\n" },
        { ControlChar::FieldEndChar, "[!FieldEnd!]" },
        { ControlChar::FieldSeparatorChar, "[!FieldSeparator!]" },
        { ControlChar::FieldStartChar, "[!FieldStart!]" },
        { ControlChar::LineBreakChar, "[!LineBreak!]\n" },
        { ControlChar::LineFeedChar, "[!LineFeed!]" },
        { ControlChar::NonBreakingHyphenChar, "[!NonBreakingHyphen!]" },
        { ControlChar::NonBreakingSpaceChar, "[!NonBreakingSpace!]" },
        { ControlChar::OptionalHyphenChar, "[!OptionalHyphen!]" },
        { ControlChar::ParagraphBreakChar, "¶\n" },
        { ControlChar::SectionBreakChar, "[!SectionBreak!]\r\n" },
        { ControlChar::TabChar, "[!Tab!]" },
    };

    auto node = mDocumentModel->getNodePointer(index);
    if (node == nullptr)
        return QString();

    QString text;
    for (char16_t c : node->GetText())
    {
        auto it = controlCharacters.find(c);
        if (it == controlCharacters.end())
        {
            text.append(QString(c).toHtmlEscaped());
        }
        else
        {
            // All control characters are converted to human readable form.
            // E.g. [!PageBreak!], [!ParagraphBreak!], etc.
            text.append("<font color=\"IndianRed\">" + it->second + "</font>");
            if (it->second.endsWith('\n'))
                text.append("<br>");
        }
    }
    return text;
}

// Selects file name for a document to open or null.
QString MainWindow::selectOpenFileName()
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

// Selects file name for saving currently opened document or null.
QString MainWindow::selectSaveFileName()
{
    const QString filter =
        "Word 97-2003 Document (*.doc);;"
        "Word 2007 OOXML Document (*.docx);;"
        "Word 2007 OOXML Macro-Enabled Document (*.docm);;"
        "PDF (*.pdf);;"
        "XPS Document (*.xps);;"
        "OpenDocument Text (*.odt);;"
        "Web Page (*.html);;"
        "Single File Web Page (*.mht);;"
        "Rich Text Format (*.rtf);;"
        "Word 2003 WordprocessingML (*.xml);;"
        "FlatOPC XML Document (*.fopc);;"
        "Plain Text (*.txt);;"
        "IDPF EPUB Document (*.epub);;"
        "XAML Fixed Document (*.xaml)";
    const QFileInfo originalFileInfo(Convert(mDocument->get_OriginalFileName()));
    const QString templateFileName = mInitialDirectory + QDir::separator() + originalFileInfo.baseName() + ".doc";
    const QString fileName = QFileDialog::getSaveFileName(this, "Save Document As", templateFileName, filter);
    if (!fileName.isNull())
    {
        const QFileInfo fileInfo(fileName);
        mInitialDirectory = fileInfo.dir().path();
    }
    return fileName;
}

// Show information about application error.
void MainWindow::showErrorMessage(QString message)
{
    QMessageBox messageBox(QMessageBox::Critical,
                           "Document Explorer - unexpected error occured",
                           message, QMessageBox::Ok, this);
    messageBox.setTextInteractionFlags(Qt::TextSelectableByMouse);
    messageBox.exec();
}

