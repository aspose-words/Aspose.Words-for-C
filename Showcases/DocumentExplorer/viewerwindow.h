#ifndef VIEWERWINDOW_H
#define VIEWERWINDOW_H

#include <QLabel>
#include <QMainWindow>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

namespace Ui {
class ViewerWindow;
}

// A simple window to show a Word document using Aspose.Words.Viewer.
class ViewerWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit ViewerWindow(System::SharedPtr<Aspose::Words::Document> document, QWidget *parent = nullptr);
    ~ViewerWindow();

private slots:
    void on_actionOpen_triggered();

    void on_actionExit_triggered();

    void on_actionPreviousPage_triggered();

    void on_actionNextPage_triggered();

    void on_actionFirstPage_triggered();

    void on_actionLastPage_triggered();

    void on_actionGoToPage_triggered();

private:
    Ui::ViewerWindow *ui;

    // The currently opened document.
    System::SharedPtr<Aspose::Words::Document> mDocument;

    // Last selected directory in the open and save dialogs.
    QString mInitialDirectory;

    // Current page.
    int mPageNumber = 0;

    void openDocument();
    void updatePage();
    QString selectOpenFileName();
    void showErrorMessage(QString message);
};

#endif // VIEWERWINDOW_H
