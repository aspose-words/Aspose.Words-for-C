#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/ImportFormatOptions.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Importing/NodeImporter.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
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

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace ApiExamples {

class ExNodeImporter : public ApiExampleBase
{
public:
    void KeepSourceNumbering()
    {
        //ExStart
        //ExFor:ImportFormatOptions.KeepSourceNumbering
        //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how the numbering will be imported when it clashes in source and destination documents.
        // Open a document with a custom list numbering scheme and clone it
        // Since both have the same numbering format, the formats will clash if we import one document into the other
        auto srcDoc = MakeObject<Document>(MyDir + u"Custom list numbering.docx");
        SharedPtr<Document> dstDoc = srcDoc->Clone();

        // Both documents have the same numbering in their lists, but if we set this flag to false and then import one document into the other
        // the numbering of the imported source document will continue from where it ends in the destination document
        auto importFormatOptions = MakeObject<ImportFormatOptions>();
        importFormatOptions->set_KeepSourceNumbering(false);

        auto importer = MakeObject<NodeImporter>(srcDoc, dstDoc, ImportFormatMode::KeepDifferentStyles, importFormatOptions);
        for (auto paragraph : System::IterateOver<Paragraph>(srcDoc->get_FirstSection()->get_Body()->get_Paragraphs()))
        {
            SharedPtr<Node> importedNode = importer->ImportNode(paragraph, true);
            dstDoc->get_FirstSection()->get_Body()->AppendChild(importedNode);
        }

        dstDoc->UpdateListLabels();
        dstDoc->Save(ArtifactsDir + u"NodeImporter.KeepSourceNumbering.docx");
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
        auto mainDoc = MakeObject<Document>(MyDir + u"Document insertion destination.docx");
        auto docToInsert = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<Bookmark> bookmark = mainDoc->get_Range()->get_Bookmarks()->idx_get(u"insertionPlace");
        InsertDocument(bookmark->get_BookmarkStart()->get_ParentNode(), docToInsert);

        mainDoc->Save(ArtifactsDir + u"NodeImporter.InsertAtBookmark.docx");
        TestInsertAtBookmark(MakeObject<Document>(ArtifactsDir + u"NodeImporter.InsertAtBookmark.docx"));
        //ExSkip
    }

private:
    /// <summary>
    /// Inserts content of the external document after the specified node.
    /// </summary>
    static void InsertDocument(SharedPtr<Node> insertionDestination, SharedPtr<Document> docToInsert)
    {
        // Make sure that the node is either a paragraph or table
        if (insertionDestination->get_NodeType() == NodeType::Paragraph || insertionDestination->get_NodeType() == NodeType::Table)
        {
            // We will be inserting into the parent of the destination paragraph
            SharedPtr<CompositeNode> dstStory = insertionDestination->get_ParentNode();

            // This object will be translating styles and lists during the import
            auto importer = MakeObject<NodeImporter>(docToInsert, insertionDestination->get_Document(), ImportFormatMode::KeepSourceFormatting);

            // Loop through all block level nodes in the body of the section
            for (auto srcSection : System::IterateOver(docToInsert->get_Sections()->LINQ_OfType<SharedPtr<Section>>()))
            {
                for (auto srcNode : System::IterateOver(srcSection->get_Body()))
                {
                    // Skip the node if it is a last empty paragraph in a section
                    if (srcNode->get_NodeType() == NodeType::Paragraph)
                    {
                        auto para = System::DynamicCast<Paragraph>(srcNode);
                        if (para->get_IsEndOfSection() && !para->get_HasChildNodes())
                        {
                            continue;
                        }
                    }

                    // This creates a clone of the node, suitable for insertion into the destination document
                    SharedPtr<Node> newNode = importer->ImportNode(srcNode, true);

                    // Insert new node after the reference node
                    dstStory->InsertAfter(newNode, insertionDestination);
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

protected:
    void TestInsertAtBookmark(SharedPtr<Document> doc)
    {
        ASSERT_EQ(String(u"1) At text that can be identified by regex:\r[MY_DOCUMENT]\r") +
                      u"2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" + u"3) At a bookmark:\r\rHello World!",
                  doc->get_FirstSection()->get_Body()->GetText().Trim());
    }

public:
    void InsertAtMailMerge()
    {
        // Open the main document
        auto mainDoc = MakeObject<Document>(MyDir + u"Document insertion destination.docx");

        // Add a handler to MergeField event
        mainDoc->get_MailMerge()->set_FieldMergingCallback(MakeObject<ExNodeImporter::InsertDocumentAtMailMergeHandler>());

        // The main document has a merge field in it called "Document_1"
        // The corresponding data for this field contains fully qualified path to the document
        // that should be inserted to this field
        mainDoc->get_MailMerge()->Execute(
            MakeArray<String>({u"Document_1"}),
            MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(MyDir + u"Document.docx")}));

        mainDoc->Save(ArtifactsDir + u"NodeImporter.InsertAtMailMerge.docx");
        TestInsertAtMailMerge(MakeObject<Document>(ArtifactsDir + u"NodeImporter.InsertAtMailMerge.docx"));
        //ExSkip
    }

private:
    class InsertDocumentAtMailMergeHandler : public IFieldMergingCallback
    {
    public:
        /// <summary>
        /// This handler makes special processing for the "Document_1" field.
        /// The field value contains the path to load the document.
        /// We load the document and insert it into the current merge field.
        /// </summary>
        void FieldMerging(SharedPtr<FieldMergingArgs> args) override
        {
            if (args->get_DocumentFieldName() == u"Document_1")
            {
                // Use document builder to navigate to the merge field with the specified name
                auto builder = MakeObject<DocumentBuilder>(args->get_Document());
                builder->MoveToMergeField(args->get_DocumentFieldName());

                // The name of the document to load and insert is stored in the field value
                auto subDoc = MakeObject<Document>(System::ObjectExt::Unbox<String>(args->get_FieldValue()));

                // Insert the document
                InsertDocument(builder->get_CurrentParagraph(), subDoc);

                // The paragraph that contained the merge field might be empty now and you probably want to delete it
                if (!builder->get_CurrentParagraph()->get_HasChildNodes())
                {
                    builder->get_CurrentParagraph()->Remove();
                }

                // Indicate to the mail merge engine that we have inserted what we wanted
                args->set_Text(nullptr);
            }
        }

        void ImageFieldMerging(SharedPtr<ImageFieldMergingArgs> args) override
        {
            // Do nothing
        }
    };

protected:
    void TestInsertAtMailMerge(SharedPtr<Document> doc)
    {
        ASSERT_EQ(String(u"1) At text that can be identified by regex:\r[MY_DOCUMENT]\r") + u"2) At a MERGEFIELD:\rHello World!\r" +
                      u"3) At a bookmark:",
                  doc->get_FirstSection()->get_Body()->GetText().Trim());
    }
};

} // namespace ApiExamples
