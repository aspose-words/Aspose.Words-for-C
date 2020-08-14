#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: System.Data is not supported

#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace MailMerging { class FieldMergingArgs; } } }
namespace Aspose { namespace Words { namespace MailMerging { class ImageFieldMergingArgs; } } }

namespace ApiExamples {

class ExMailMergeEvent : public ApiExampleBase
{
private:

    class HandleMergeFieldInsertHtml : public Aspose::Words::MailMerging::IFieldMergingCallback
    {
        typedef HandleMergeFieldInsertHtml ThisType;
        typedef Aspose::Words::MailMerging::IFieldMergingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// This is called when merge field is actually merged with data in the document.
        /// </summary>
        void FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args) override;
        void ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args) override;
        
    };
    
    
public:

    void InsertHtml();
    void ImageFromUrl();
    
};

} // namespace ApiExamples


