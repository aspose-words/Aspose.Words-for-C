#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Bookmark.h>
#include <Aspose.Words.Cpp/BookmarkCollection.h>
#include <Aspose.Words.Cpp/BookmarkEnd.h>
#include <Aspose.Words.Cpp/BookmarkStart.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/ImportFormatMode.h>
#include <Aspose.Words.Cpp/ImportFormatOptions.h>
#include <Aspose.Words.Cpp/MailMerging/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/MailMerging/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/MailMerging/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeImporter.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace ApiExamples {

class ExNodeImporter : public ApiExampleBase
{
public:
    void KeepSourceNumbering(bool keepSourceNumbering)
    {
        //ExStart
        //ExFor:ImportFormatOptions.KeepSourceNumbering
        //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to resolve list numbering clashes in source and destination documents.
        // Open a document with a custom list numbering scheme, and then clone it.
        // Since both have the same numbering format, the formats will clash if we import one document into the other.
        auto srcDoc = MakeObject<Document>(MyDir + u"Custom list numbering.docx");
        SharedPtr<Document> dstDoc = srcDoc->Clone();

        // When we import the document's clone into the original and then append it,
        // then the two lists with the same list format will join.
        // If we set the "KeepSourceNumbering" flag to "false", then the list from the document clone
        // that we append to the original will carry on the numbering of the list we append it to.
        // This will effectively merge the two lists into one.
        // If we set the "KeepSourceNumbering" flag to "true", then the document clone
        // list will preserve its original numbering, making the two lists appear as separate lists.
        auto importFormatOptions = MakeObject<ImportFormatOptions>();
        importFormatOptions->set_KeepSourceNumbering(keepSourceNumbering);

        auto importer = MakeObject<NodeImporter>(srcDoc, dstDoc, ImportFormatMode::KeepDifferentStyles, importFormatOptions);
        for (const auto& paragraph : System::IterateOver<Paragraph>(srcDoc->get_FirstSection()->get_Body()->get_Paragraphs()))
        {
            SharedPtr<Node> importedNode = importer->ImportNode(paragraph, true);
            dstDoc->get_FirstSection()->get_Body()->AppendChild(importedNode);
        }

        dstDoc->UpdateListLabels();

        if (keepSourceNumbering)
        {
            ASSERT_EQ(String(u"6. Item 1\r\n") + u"7. Item 2 \r\n" + u"8. Item 3\r\n" + u"9. Item 4\r\n" + u"6. Item 1\r\n" + u"7. Item 2 \r\n" +
                          u"8. Item 3\r\n" + u"9. Item 4",
                      dstDoc->get_FirstSection()->get_Body()->ToString(SaveFormat::Text).Trim());
        }
        else
        {
            ASSERT_EQ(String(u"6. Item 1\r\n") + u"7. Item 2 \r\n" + u"8. Item 3\r\n" + u"9. Item 4\r\n" + u"10. Item 1\r\n" + u"11. Item 2 \r\n" +
                          u"12. Item 3\r\n" + u"13. Item 4",
                      dstDoc->get_FirstSection()->get_Body()->ToString(SaveFormat::Text).Trim());
        }
        //ExEnd
    }

    //ExStart
    //ExFor:Paragraph.IsEndOfSection
    //ExFor:NodeImporter
    //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode)
    //ExFor:NodeImporter.ImportNode(Node, Boolean)
    //ExSummary:Shows how to insert the contents of one document to a bookmark in another document.
    void InsertAtBookmark()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartBookmark(u"InsertionPoint");
        builder->Write(u"We will insert a document here: ");
        builder->EndBookmark(u"InsertionPoint");

        auto docToInsert = MakeObject<Document>();
        builder = MakeObject<DocumentBuilder>(docToInsert);

        builder->Write(u"Hello world!");

        docToInsert->Save(ArtifactsDir + u"NodeImporter.InsertAtMergeField.docx");

        SharedPtr<Bookmark> bookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"InsertionPoint");
        InsertDocument(bookmark->get_BookmarkStart()->get_ParentNode(), docToInsert);

        ASSERT_EQ(String(u"We will insert a document here: ") + u"\rHello world!", doc->GetText().Trim());
    }

    /// <summary>
    /// Inserts the contents of a document after the specified node.
    /// </summary>
    static void InsertDocument(SharedPtr<Node> insertionDestination, SharedPtr<Document> docToInsert)
    {
        if (insertionDestination->get_NodeType() == NodeType::Paragraph || insertionDestination->get_NodeType() == NodeType::Table)
        {
            SharedPtr<CompositeNode> destinationParent = insertionDestination->get_ParentNode();

            auto importer = MakeObject<NodeImporter>(docToInsert, insertionDestination->get_Document(), ImportFormatMode::KeepSourceFormatting);

            // Loop through all block-level nodes in the section's body,
            // then clone and insert every node that is not the last empty paragraph of a section.
            for (const auto& srcSection : System::IterateOver(docToInsert->get_Sections()->LINQ_OfType<SharedPtr<Section>>()))
            {
                for (const auto& srcNode : System::IterateOver(srcSection->get_Body()))
                {
                    if (srcNode->get_NodeType() == NodeType::Paragraph)
                    {
                        auto para = System::DynamicCast<Paragraph>(srcNode);
                        if (para->get_IsEndOfSection() && !para->get_HasChildNodes())
                        {
                            continue;
                        }
                    }

                    SharedPtr<Node> newNode = importer->ImportNode(srcNode, true);

                    destinationParent->InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
        else
        {
            throw System::ArgumentException(u"The destination node should be either a paragraph or table.");
        }
    }
    //ExEnd

    void InsertAtMergeField()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u"A document will appear here: ");
        builder->InsertField(u" MERGEFIELD Document_1 ");

        auto subDoc = MakeObject<Document>();
        builder = MakeObject<DocumentBuilder>(subDoc);
        builder->Write(u"Hello world!");

        subDoc->Save(ArtifactsDir + u"NodeImporter.InsertAtMergeField.docx");

        doc->get_MailMerge()->set_FieldMergingCallback(MakeObject<ExNodeImporter::InsertDocumentAtMailMergeHandler>());

        // The main document has a merge field in it called "Document_1".
        // Execute a mail merge using a data source that contains a local system filename
        // of the document that we wish to insert into the MERGEFIELD.
        doc->get_MailMerge()->Execute(
            MakeArray<String>({u"Document_1"}),
            MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(ArtifactsDir + u"NodeImporter.InsertAtMergeField.docx")}));

        ASSERT_EQ(String(u"A document will appear here: \r") + u"Hello world!", doc->GetText().Trim());
    }

    /// <summary>
    /// If the mail merge encounters a MERGEFIELD with a specified name,
    /// this handler treats the current value of a mail merge data source as a local system filename of a document.
    /// The handler will insert the document in its entirety into the MERGEFIELD instead of the current merge value.
    /// </summary>
    class InsertDocumentAtMailMergeHandler : public IFieldMergingCallback
    {
    public:
        void FieldMerging(SharedPtr<FieldMergingArgs> args) override
        {
            if (args->get_DocumentFieldName() == u"Document_1")
            {
                auto builder = MakeObject<DocumentBuilder>(args->get_Document());
                builder->MoveToMergeField(args->get_DocumentFieldName());

                auto subDoc = MakeObject<Document>(System::ObjectExt::Unbox<String>(args->get_FieldValue()));

                InsertDocument(builder->get_CurrentParagraph(), subDoc);

                if (!builder->get_CurrentParagraph()->get_HasChildNodes())
                {
                    builder->get_CurrentParagraph()->Remove();
                }

                args->set_Text(nullptr);
            }
        }

        void ImageFieldMerging(SharedPtr<ImageFieldMergingArgs> args) override
        {
            // Do nothing.
        }
    };
};

} // namespace ApiExamples
