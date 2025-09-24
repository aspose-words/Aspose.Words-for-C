#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/shared_ptr.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::MailMerging;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExNodeImporter : public ApiExampleBase
{
    typedef ExNodeImporter ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// If the mail merge encounters a MERGEFIELD with a specified name,
    /// this handler treats the current value of a mail merge data source as a local system filename of a document.
    /// The handler will insert the document in its entirety into the MERGEFIELD instead of the current merge value.
    /// </summary>
    class InsertDocumentAtMailMergeHandler : public IFieldMergingCallback
    {
        typedef InsertDocumentAtMailMergeHandler ThisType;
        typedef IFieldMergingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    private:
    
        void FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args) override;
        void ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args) override;
        
    };
    
    
public:

    void KeepSourceNumbering(bool keepSourceNumbering);
    void InsertAtBookmark();
    void InsertAtMergeField();
    
private:

    /// <summary>
    /// Inserts the contents of a document after the specified node.
    /// </summary>
    static void InsertDocument(System::SharedPtr<Aspose::Words::Node> insertionDestination, System::SharedPtr<Aspose::Words::Document> docToInsert);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


