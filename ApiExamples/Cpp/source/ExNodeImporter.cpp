// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExNodeImporter.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/Importing/NodeImporter.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(4037147130u, ::ApiExamples::ExNodeImporter::InsertDocumentAtMailMergeHandler, ThisTypeBaseTypesInfo);

void ExNodeImporter::InsertDocumentAtMailMergeHandler::FieldMerging(SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args)
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

void ExNodeImporter::InsertDocumentAtMailMergeHandler::ImageFieldMerging(SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args)
{
    // Do nothing
}

void ExNodeImporter::InsertDocument(SharedPtr<Aspose::Words::Node> insertionDestination, SharedPtr<Aspose::Words::Document> docToInsert)
{
    // Make sure that the node is either a paragraph or table
    if (System::ObjectExt::Equals(insertionDestination->get_NodeType(), Aspose::Words::NodeType::Paragraph) || System::ObjectExt::Equals(insertionDestination->get_NodeType(), Aspose::Words::NodeType::Table))
    {
        // We will be inserting into the parent of the destination paragraph
        SharedPtr<CompositeNode> dstStory = insertionDestination->get_ParentNode();

        // This object will be translating styles and lists during the import
        auto importer = MakeObject<NodeImporter>(docToInsert, insertionDestination->get_Document(), Aspose::Words::ImportFormatMode::KeepSourceFormatting);

        // Loop through all block level nodes in the body of the section
        for (auto srcSection : System::IterateOver(docToInsert->get_Sections()->LINQ_OfType<SharedPtr<Section> >()))
        {
            for (auto srcNode : System::IterateOver(srcSection->get_Body()))
            {
                // Let's skip the node if it is a last empty paragraph in a section
                if (System::ObjectExt::Equals(srcNode->get_NodeType(), Aspose::Words::NodeType::Paragraph))
                {
                    auto para = System::DynamicCast<Aspose::Words::Paragraph>(srcNode);
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

void ExNodeImporter::TestInsertAtBookmark(SharedPtr<Aspose::Words::Document> doc)
{
    ASSERT_EQ(String(u"1) At text that can be identified by regex:\r[MY_DOCUMENT]\r") + u"2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" + u"3) At a bookmark:\r\rHello World!", doc->get_FirstSection()->get_Body()->GetText().Trim());
}

void ExNodeImporter::TestInsertAtMailMerge(SharedPtr<Aspose::Words::Document> doc)
{
    ASSERT_EQ(String(u"1) At text that can be identified by regex:\r[MY_DOCUMENT]\r") + u"2) At a MERGEFIELD:\rHello World!\r" + u"3) At a bookmark:", doc->get_FirstSection()->get_Body()->GetText().Trim());
}

namespace gtest_test
{

class ExNodeImporter : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExNodeImporter> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExNodeImporter>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExNodeImporter> ExNodeImporter::s_instance;

} // namespace gtest_test

void ExNodeImporter::InsertAtBookmark()
{
    auto mainDoc = MakeObject<Document>(MyDir + u"Document insertion destination.docx");
    auto docToInsert = MakeObject<Document>(MyDir + u"Document.docx");

    SharedPtr<Bookmark> bookmark = mainDoc->get_Range()->get_Bookmarks()->idx_get(u"insertionPlace");
    InsertDocument(bookmark->get_BookmarkStart()->get_ParentNode(), docToInsert);

    mainDoc->Save(ArtifactsDir + u"InsertDocument.InsertAtBookmark.docx");
    TestInsertAtBookmark(MakeObject<Document>(ArtifactsDir + u"InsertDocument.InsertAtBookmark.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExNodeImporter, InsertAtBookmark)
{
    s_instance->InsertAtBookmark();
}

} // namespace gtest_test

void ExNodeImporter::InsertAtMailMerge()
{
    // Open the main document
    auto mainDoc = MakeObject<Document>(MyDir + u"Document insertion destination.docx");

    // Add a handler to MergeField event
    mainDoc->get_MailMerge()->set_FieldMergingCallback(MakeObject<ExNodeImporter::InsertDocumentAtMailMergeHandler>());

    // The main document has a merge field in it called "Document_1"
    // The corresponding data for this field contains fully qualified path to the document
    // that should be inserted to this field
    mainDoc->get_MailMerge()->Execute(MakeArray<String>({u"Document_1"}), MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(MyDir + u"Document.docx")}));

    mainDoc->Save(ArtifactsDir + u"InsertDocument.InsertAtMailMerge.docx");
    TestInsertAtMailMerge(MakeObject<Document>(ArtifactsDir + u"InsertDocument.InsertAtMailMerge.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExNodeImporter, InsertAtMailMerge)
{
    s_instance->InsertAtMailMerge();
}

} // namespace gtest_test

} // namespace ApiExamples
