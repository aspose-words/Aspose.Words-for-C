#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/shared_ptr.h>
#include <system/io/stream.h>
#include <Aspose.Words.Cpp/Model/Fonts/StreamFontSource.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Fonts;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExFontSettings : public ApiExampleBase
{
    typedef ExFontSettings ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    class HandleDocumentWarnings : public IWarningCallback
    {
        typedef HandleDocumentWarnings ThisType;
        typedef IWarningCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<Aspose::Words::WarningInfoCollection> FontWarnings;
        
        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        void Warning(System::SharedPtr<Aspose::Words::WarningInfo> info) override;
        
        HandleDocumentWarnings();
        
    };
    
    
private:

    class FontSubstitutionWarningCollector : public IWarningCallback
    {
        typedef FontSubstitutionWarningCollector ThisType;
        typedef IWarningCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<Aspose::Words::WarningInfoCollection> FontSubstitutionWarnings;
        
        /// <summary>
        /// Called every time a warning occurs during loading/saving.
        /// </summary>
        void Warning(System::SharedPtr<Aspose::Words::WarningInfo> info) override;
        
        FontSubstitutionWarningCollector();
        
    };
    
    class FontSourceWarningCollector : public IWarningCallback
    {
        typedef FontSourceWarningCollector ThisType;
        typedef IWarningCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<Aspose::Words::WarningInfoCollection> FontSubstitutionWarnings;
        
        /// <summary>
        /// Called every time a warning occurs during processing of font source.
        /// </summary>
        void Warning(System::SharedPtr<Aspose::Words::WarningInfo> info) override;
        
        FontSourceWarningCollector();
        
    };
    
    /// <summary>
    /// Load the font data only when required instead of storing it in the memory
    /// for the entire lifetime of the "FontSettings" object.
    /// </summary>
    class StreamFontSourceFile : public StreamFontSource
    {
        typedef StreamFontSourceFile ThisType;
        typedef StreamFontSource BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<System::IO::Stream> OpenFontDataStream() override;
        
    protected:
    
        virtual ~StreamFontSourceFile();
        
    };
    
    
public:

    void DefaultFontInstance();
    void DefaultFontName();
    void UpdatePageLayoutWarnings();
    void SubstitutionWarning();
    void FontSourceWarning();
    void EnableFontSubstitution();
    void SubstitutionWarningsClosestMatch();
    void DisableFontSubstitution();
    void SubstitutionWarnings();
    void GetSubstitutionWithoutSuffixes();
    void FontSourceFile();
    void FontSourceFolder();
    void SetFontsFolder(bool recursive);
    void SetFontsFolders(bool recursive);
    void AddFontSource();
    void SetSpecifyFontFolder();
    void TableSubstitution();
    void SetSpecifyFontFolders();
    void AddFontSubstitutes();
    void FontSourceMemory();
    void FontSourceSystem();
    void LoadFontFallbackSettingsFromFile();
    void LoadFontFallbackSettingsFromStream();
    void LoadNotoFontsFallbackSettings();
    void DefaultFontSubstitutionRule();
    void FontConfigSubstitution();
    void FallbackSettings();
    void FallbackSettingsCustom();
    void TableSubstitutionRule();
    void TableSubstitutionRuleCustom();
    void ResolveFontsBeforeLoadingDocument();
    void StreamFontSourceFileRendering();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


