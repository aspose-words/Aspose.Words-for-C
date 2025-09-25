#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/shared_ptr.h>
#include <system/date_time.h>
#include <system/collections/list.h>
#include <Aspose.Words.Cpp/Model/Progress/IDocumentLoadingCallback.h>
#include <Aspose.Words.Cpp/Model/Progress/DocumentLoadingArgs.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingArgs.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingAction.h>
#include <Aspose.Words.Cpp/Model/Loading/IResourceLoadingCallback.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Loading;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExLoadOptions : public ApiExampleBase
{
    typedef ExLoadOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Cancel a document loading after the "MaxDuration" seconds.
    /// </summary>
    class LoadingProgressCallback : public IDocumentLoadingCallback
    {
        typedef LoadingProgressCallback ThisType;
        typedef IDocumentLoadingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// Ctr.
        /// </summary>
        LoadingProgressCallback();
        
        /// <summary>
        /// Callback method which called during document loading.
        /// </summary>
        /// <param name="args">Loading arguments.</param>
        void Notify(System::SharedPtr<Aspose::Words::Loading::DocumentLoadingArgs> args) override;
        
    private:
    
        /// <summary>
        /// Date and time when document loading is started.
        /// </summary>
        System::DateTime mLoadingStartedAt;
        /// <summary>
        /// Maximum allowed duration in sec.
        /// </summary>
        static const double MaxDuration;
        
    };
    
    
private:

    /// <summary>
    /// Prints the filenames of all external stylesheets and substitutes all images of a loaded html document.
    /// </summary>
    class HtmlLinkedResourceLoadingCallback : public IResourceLoadingCallback
    {
        typedef HtmlLinkedResourceLoadingCallback ThisType;
        typedef IResourceLoadingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        Aspose::Words::Loading::ResourceLoadingAction ResourceLoading(System::SharedPtr<Aspose::Words::Loading::ResourceLoadingArgs> args) override;
        
    };
    
    /// <summary>
    /// IWarningCallback that prints warnings and their details as they arise during document loading.
    /// </summary>
    class DocumentLoadingWarningCallback : public IWarningCallback
    {
        typedef DocumentLoadingWarningCallback ThisType;
        typedef IWarningCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void Warning(System::SharedPtr<Aspose::Words::WarningInfo> info) override;
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> GetWarnings();
        
        DocumentLoadingWarningCallback();
        
    private:
    
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> mWarnings;
        
    };
    
    
public:

    void LoadOptionsCallback();
    void ConvertShapeToOfficeMath(bool isConvertShapeToOfficeMath);
    void SetEncoding();
    void FontSettings();
    void LoadOptionsMswVersion();
    void LoadOptionsWarningCallback();
    void TempFolder();
    void AddEditingLanguage();
    void SetEditingLanguageAsDefault();
    void ConvertMetafilesToPng();
    void OpenChmFile();
    void ProgressCallback();
    void IgnoreOleData();
    void RecoveryMode();
    
protected:

    static void TestLoadOptionsWarningCallback(System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> warnings);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


