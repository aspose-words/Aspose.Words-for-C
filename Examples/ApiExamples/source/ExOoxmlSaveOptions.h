#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/date_time.h>
#include <Aspose.Words.Cpp/Model/Saving/CompressionLevel.h>
#include <Aspose.Words.Cpp/Model/Progress/IDocumentSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Progress/DocumentSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Saving;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExOoxmlSaveOptions : public ApiExampleBase
{
    typedef ExOoxmlSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Saving progress callback. Cancel a document saving after the "MaxDuration" seconds.
    /// </summary>
    class SavingProgressCallback : public IDocumentSavingCallback
    {
        typedef SavingProgressCallback ThisType;
        typedef IDocumentSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// Ctr.
        /// </summary>
        SavingProgressCallback();
        
        /// <summary>
        /// Callback method which called during document saving.
        /// </summary>
        /// <param name="args">Saving arguments.</param>
        void Notify(System::SharedPtr<Aspose::Words::Saving::DocumentSavingArgs> args) override;
        
    private:
    
        /// <summary>
        /// Date and time when document saving is started.
        /// </summary>
        System::DateTime mSavingStartedAt;
        /// <summary>
        /// Maximum allowed duration in sec.
        /// </summary>
        static const double MaxDuration;
        
    };
    
    
public:

    void Password();
    void Iso29500Strict();
    void RestartingDocumentList(bool restartListAtEachSection);
    void LastSavedTime(bool updateLastSavedTimeProperty);
    void KeepLegacyControlChars(bool keepLegacyControlChars);
    void DocumentCompression(Aspose::Words::Saving::CompressionLevel compressionLevel);
    void CheckFileSignatures();
    void ExportGeneratorName();
    void ProgressCallback(Aspose::Words::SaveFormat saveFormat, System::String ext);
    void Zip64ModeOption();
    void DigitalSignature();
    void UpdateAmbiguousTextFont();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


