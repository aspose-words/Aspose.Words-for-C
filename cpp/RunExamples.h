#pragma once

#include <iostream>
#include <system/string.h>

System::String GetDataDir_QuickStart();

System::String GetOutputFilePath(const System::String& inputFilePath);

System::String GetDataDir_LINQ();

System::String GetDataDir_Database();

System::String GetDataDir_LoadingAndSaving();

System::String GetDataDir_JoiningAndAppending();

System::String GetDataDir_WorkingWithList();

System::String GetDataDir_FindAndReplace();

System::String GetDataDir_ConvertUtil();

System::String GetDataDir_WorkingWithBookmarks();

System::String GetDataDir_WorkingWithComments();

System::String GetDataDir_WorkingWithDocument();

System::String GetDataDir_WorkingWithShapes();

System::String GetDataDir_WorkingWithFields();

System::String GetDataDir_WorkingWithHyperlink();

System::String GetDataDir_WorkingWithCharts();

System::String GetDataDir_WorkingWithOnlineVideo();

System::String GetDataDir_WorkingWithNode();

System::String GetDataDir_WorkingWithTheme();

System::String GetDataDir_WorkingWithRanges();

System::String GetDataDir_WorkingWithStructuredDocumentTag();

System::String GetDataDir_WorkingWithImages();

System::String GetDataDir_WorkingWithStyles();

System::String GetDataDir_WorkingWithTables();

System::String GetDataDir_WorkingWithSignature();

System::String GetDataDir_WorkingWithSections();

System::String GetDataDir_MailMergeAndReporting();

System::String GetDataDir_RenderingAndPrinting();

System::String GetDataDir_ViewersAndVisualizers();

template <typename T> struct DisposableHolder
{
public:
    DisposableHolder(System::SharedPtr<T> obj) : mObj(obj)
    {
    }
    DisposableHolder(DisposableHolder<T>&) = delete;
    DisposableHolder(DisposableHolder<T>&&) = delete;
    DisposableHolder<T> operator=(DisposableHolder<T>&) = delete;
    DisposableHolder<T> operator=(DisposableHolder<T>&&) = delete;
    ~DisposableHolder()
    {
        try
        {
            mObj->Dispose();
        }
        catch (System::Exception &exc)
        {
            std::cerr << "The following error is raised during the call of the Dispose method: " << exc.ToString().ToUtf8String() << std::endl;
        }
        catch(...)
        {
            std::cerr << "The unknown error is raised during the call of the Dispose method" << std::endl;
        }
    }

    System::SharedPtr<T> GetObject() const { return mObj; }

private:
    System::SharedPtr<T> mObj;
};