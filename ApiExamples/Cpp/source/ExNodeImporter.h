#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class Node; } }
namespace Aspose { namespace Words { class Document; } }
namespace Aspose { namespace Words { namespace MailMerging { class FieldMergingArgs; } } }
namespace Aspose { namespace Words { namespace MailMerging { class ImageFieldMergingArgs; } } }

namespace ApiExamples {

class ExNodeImporter : public ApiExampleBase
{
private:

    class InsertDocumentAtMailMergeHandler : public Aspose::Words::MailMerging::IFieldMergingCallback
    {
        typedef InsertDocumentAtMailMergeHandler ThisType;
        typedef Aspose::Words::MailMerging::IFieldMergingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// This handler makes special processing for the "Document_1" field.
        /// The field value contains the path to load the document. 
        /// We load the document and insert it into the current merge field.
        /// </summary>
        void FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args) override;
        void ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args) override;
        
    };
    
    
public:

    void InsertAtBookmark();
    void InsertAtMailMerge();
    
protected:

    void TestInsertAtBookmark(System::SharedPtr<Aspose::Words::Document> doc);
    void TestInsertAtMailMerge(System::SharedPtr<Aspose::Words::Document> doc);
    
private:

    /// <summary>
    /// Inserts content of the external document after the specified node.
    /// </summary>
    static void InsertDocument(System::SharedPtr<Aspose::Words::Node> insertionDestination, System::SharedPtr<Aspose::Words::Document> docToInsert);
    
};

} // namespace ApiExamples


