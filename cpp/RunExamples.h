#pragma once

#include <iostream>
#include <system/string.h>

System::String GetInputDataDir_QuickStart();

System::String GetOutputDataDir_QuickStart();

System::String GetInputDataDir_LINQ();

System::String GetOutputDataDir_LINQ();

System::String GetInputDataDir_Database();

System::String GetOutputDataDir_Database();

System::String GetInputDataDir_LoadingAndSaving();

System::String GetOutputDataDir_LoadingAndSaving();

System::String GetInputDataDir_JoiningAndAppending();

System::String GetOutputDataDir_JoiningAndAppending();

System::String GetInputDataDir_WorkingWithList();

System::String GetOutputDataDir_WorkingWithList();

System::String GetInputDataDir_FindAndReplace();

System::String GetOutputDataDir_FindAndReplace();

System::String GetInputDataDir_ConvertUtil();

System::String GetOutputDataDir_ConvertUtil();

System::String GetInputDataDir_WorkingWithBookmarks();

System::String GetOutputDataDir_WorkingWithBookmarks();

System::String GetInputDataDir_WorkingWithComments();

System::String GetOutputDataDir_WorkingWithComments();

System::String GetInputDataDir_WorkingWithDocument();

System::String GetOutputDataDir_WorkingWithDocument();

System::String GetInputDataDir_WorkingWithShapes();

System::String GetOutputDataDir_WorkingWithShapes();

System::String GetInputDataDir_WorkingWithFields();

System::String GetOutputDataDir_WorkingWithFields();

System::String GetInputDataDir_WorkingWithHyperlink();

System::String GetOutputDataDir_WorkingWithHyperlink();

System::String GetInputDataDir_WorkingWithCharts();

System::String GetOutputDataDir_WorkingWithCharts();

System::String GetInputDataDir_WorkingWithOnlineVideo();

System::String GetOutputDataDir_WorkingWithOnlineVideo();

System::String GetInputDataDir_WorkingWithNode();

System::String GetOutputDataDir_WorkingWithNode();

System::String GetInputDataDir_WorkingWithTheme();

System::String GetOutputDataDir_WorkingWithTheme();

System::String GetInputDataDir_WorkingWithRanges();

System::String GetOutputDataDir_WorkingWithRanges();

System::String GetInputDataDir_WorkingWithStructuredDocumentTag();

System::String GetOutputDataDir_WorkingWithStructuredDocumentTag();

System::String GetInputDataDir_WorkingWithImages();

System::String GetOutputDataDir_WorkingWithImages();

System::String GetInputDataDir_WorkingWithStyles();

System::String GetOutputDataDir_WorkingWithStyles();

System::String GetInputDataDir_WorkingWithTables();

System::String GetOutputDataDir_WorkingWithTables();

System::String GetInputDataDir_WorkingWithSignature();

System::String GetOutputDataDir_WorkingWithSignature();

System::String GetInputDataDir_WorkingWithSections();

System::String GetOutputDataDir_WorkingWithSections();

System::String GetInputDataDir_MailMergeAndReporting();

System::String GetOutputDataDir_MailMergeAndReporting();

System::String GetInputDataDir_RenderingAndPrinting();

System::String GetOutputDataDir_RenderingAndPrinting();

System::String GetInputDataDir_ViewersAndVisualizers();

System::String GetOutputDataDir_ViewersAndVisualizers();

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