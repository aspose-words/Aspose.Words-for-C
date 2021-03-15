#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <memory>
#include "documentmodel.h"

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

// The main window of the DocumentExplorer demo.
//
// DocumentExplorer allows to open documents using Aspose.Words.
// Once a document is opened, you can explore its object model in the tree.
// You can also save the document into DOC, DOCX, ODF, EPUB, PDF, RTF, WordML,
// HTML, MHTML and plain text formats.
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void on_actionOpen_triggered();

    void on_actionSaveAs_triggered();

    void on_actionRender_triggered();

    void on_actionExit_triggered();

    void on_actionRemoveNode_triggered();

    void on_actionExpandAll_triggered();

    void on_actionCollapseAll_triggered();

    void on_actionAbout_triggered();

    void treeView_currentChanged(const QModelIndex &current, const QModelIndex &previous);

private:
    Ui::MainWindow *ui;

    // The currently opened document.
    System::SharedPtr<Aspose::Words::Document> mDocument;

    // Hierarchical model of the opened document.
    std::unique_ptr<DocumentModel> mDocumentModel;

    // Last selected directory in the open and save dialogs.
    QString mInitialDirectory;

    static void findAndApplyLicense();
    static void licenseAsposeWords(QString licenseFile);

    void openDocument();
    void saveDocument();
    void renderDocument();
    void removeNode();
    void expandAllNodes();
    void collapseAllNodes();
    void collapseNodesRecursively(const QModelIndex& parent);
    QString getNodeText(const QModelIndex& index) const;
    QString selectOpenFileName();
    QString selectSaveFileName();
    void showErrorMessage(QString message);
};
#endif // MAINWINDOW_H
