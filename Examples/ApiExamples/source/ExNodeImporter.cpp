// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExNodeImporter.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/Importing/NodeImporter.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/ImportFormatOptions.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>


using namespace Aspose::Words::MailMerging;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2191283764u, ::Aspose::Words::ApiExamples::ExNodeImporter::InsertDocumentAtMailMergeHandler, ThisTypeBaseTypesInfo);

void ExNodeImporter::InsertDocumentAtMailMergeHandler::FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args)
{
    if (args->get_DocumentFieldName() == u"Document_1")
    {
        auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(args->get_Document());
        builder->MoveToMergeField(args->get_DocumentFieldName());
        
        auto subDoc = System::MakeObject<Aspose::Words::Document>(System::ExplicitCast<System::String>(args->get_FieldValue()));
        
        InsertDocument(builder->get_CurrentParagraph(), subDoc);
        
        if (!builder->get_CurrentParagraph()->get_HasChildNodes())
        {
            builder->get_CurrentParagraph()->Remove();
        }
        
        args->set_Text(nullptr);
    }
}

void ExNodeImporter::InsertDocumentAtMailMergeHandler::ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args)
{
    // Do nothing.
}


RTTI_INFO_IMPL_HASH(1557908164u, ::Aspose::Words::ApiExamples::ExNodeImporter, ThisTypeBaseTypesInfo);

void ExNodeImporter::InsertDocument(System::SharedPtr<Aspose::Words::Node> insertionDestination, System::SharedPtr<Aspose::Words::Document> docToInsert)
{
    if (insertionDestination->get_NodeType() == Aspose::Words::NodeType::Paragraph || insertionDestination->get_NodeType() == Aspose::Words::NodeType::Table)
    {
        System::SharedPtr<Aspose::Words::CompositeNode> destinationParent = insertionDestination->get_ParentNode();
        
        auto importer = System::MakeObject<Aspose::Words::NodeImporter>(docToInsert, insertionDestination->get_Document(), Aspose::Words::ImportFormatMode::KeepSourceFormatting);
        
        // Loop through all block-level nodes in the section's body,
        // then clone and insert every node that is not the last empty paragraph of a section.
        for (auto&& srcSection : System::IterateOver(docToInsert->get_Sections()->LINQ_OfType<System::SharedPtr<Aspose::Words::Section> >()))
        {
            for (auto&& srcNode : System::IterateOver(srcSection->get_Body()))
            {
                if (srcNode->get_NodeType() == Aspose::Words::NodeType::Paragraph)
                {
                    auto para = System::ExplicitCast<Aspose::Words::Paragraph>(srcNode);
                    if (para->get_IsEndOfSection() && !para->get_HasChildNodes())
                    {
                        continue;
                    }
                }
                
                System::SharedPtr<Aspose::Words::Node> newNode = importer->ImportNode(srcNode, true);
                
                destinationParent->InsertAfter<System::SharedPtr<Aspose::Words::Node>>(newNode, insertionDestination);
                insertionDestination = newNode;
            }
        }
    }
    else
    {
        throw System::ArgumentException(u"The destination node should be either a paragraph or table.");
    }
}


namespace gtest_test
{

class ExNodeImporter : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExNodeImporter> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExNodeImporter>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExNodeImporter> ExNodeImporter::s_instance;

} // namespace gtest_test

void ExNodeImporter::KeepSourceNumbering(bool keepSourceNumbering)
{
    //ExStart
    //ExFor:ImportFormatOptions.KeepSourceNumbering
    //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to resolve list numbering clashes in source and destination documents.
    // Open a document with a custom list numbering scheme, and then clone it.
    // Since both have the same numbering format, the formats will clash if we import one document into the other.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Custom list numbering.docx");
    System::SharedPtr<Aspose::Words::Document> dstDoc = srcDoc->Clone();
    
    // When we import the document's clone into the original and then append it,
    // then the two lists with the same list format will join.
    // If we set the "KeepSourceNumbering" flag to "false", then the list from the document clone
    // that we append to the original will carry on the numbering of the list we append it to.
    // This will effectively merge the two lists into one.
    // If we set the "KeepSourceNumbering" flag to "true", then the document clone
    // list will preserve its original numbering, making the two lists appear as separate lists. 
    auto importFormatOptions = System::MakeObject<Aspose::Words::ImportFormatOptions>();
    importFormatOptions->set_KeepSourceNumbering(keepSourceNumbering);
    
    auto importer = System::MakeObject<Aspose::Words::NodeImporter>(srcDoc, dstDoc, Aspose::Words::ImportFormatMode::KeepDifferentStyles, importFormatOptions);
    for (auto&& paragraph : System::IterateOver<Aspose::Words::Paragraph>(srcDoc->get_FirstSection()->get_Body()->get_Paragraphs()))
    {
        System::SharedPtr<Aspose::Words::Node> importedNode = importer->ImportNode(paragraph, true);
        dstDoc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Node>>(importedNode);
    }
    
    dstDoc->UpdateListLabels();
    
    if (keepSourceNumbering)
    {
        ASSERT_EQ(System::String(u"6. Item 1\r\n") + u"7. Item 2 \r\n" + u"8. Item 3\r\n" + u"9. Item 4\r\n" + u"6. Item 1\r\n" + u"7. Item 2 \r\n" + u"8. Item 3\r\n" + u"9. Item 4", dstDoc->get_FirstSection()->get_Body()->ToString(Aspose::Words::SaveFormat::Text).Trim());
    }
    else
    {
        ASSERT_EQ(System::String(u"6. Item 1\r\n") + u"7. Item 2 \r\n" + u"8. Item 3\r\n" + u"9. Item 4\r\n" + u"10. Item 1\r\n" + u"11. Item 2 \r\n" + u"12. Item 3\r\n" + u"13. Item 4", dstDoc->get_FirstSection()->get_Body()->ToString(Aspose::Words::SaveFormat::Text).Trim());
    }
    //ExEnd
}

namespace gtest_test
{

using ExNodeImporter_KeepSourceNumbering_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExNodeImporter::KeepSourceNumbering)>::type;

struct ExNodeImporter_KeepSourceNumbering : public ExNodeImporter, public Aspose::Words::ApiExamples::ExNodeImporter, public ::testing::WithParamInterface<ExNodeImporter_KeepSourceNumbering_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExNodeImporter_KeepSourceNumbering, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->KeepSourceNumbering(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExNodeImporter_KeepSourceNumbering, ::testing::ValuesIn(ExNodeImporter_KeepSourceNumbering::TestCases()));

} // namespace gtest_test

void ExNodeImporter::InsertAtBookmark()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->StartBookmark(u"InsertionPoint");
    builder->Write(u"We will insert a document here: ");
    builder->EndBookmark(u"InsertionPoint");
    
    auto docToInsert = System::MakeObject<Aspose::Words::Document>();
    builder = System::MakeObject<Aspose::Words::DocumentBuilder>(docToInsert);
    
    builder->Write(u"Hello world!");
    
    docToInsert->Save(get_ArtifactsDir() + u"NodeImporter.InsertAtMergeField.docx");
    
    System::SharedPtr<Aspose::Words::Bookmark> bookmark = doc->get_Range()->get_Bookmarks()->idx_get(u"InsertionPoint");
    InsertDocument(bookmark->get_BookmarkStart()->get_ParentNode(), docToInsert);
    
    ASSERT_EQ(System::String(u"We will insert a document here: ") + u"\rHello world!", doc->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExNodeImporter, InsertAtBookmark)
{
    s_instance->InsertAtBookmark();
}

} // namespace gtest_test

void ExNodeImporter::InsertAtMergeField()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Write(u"A document will appear here: ");
    builder->InsertField(u" MERGEFIELD Document_1 ");
    
    auto subDoc = System::MakeObject<Aspose::Words::Document>();
    builder = System::MakeObject<Aspose::Words::DocumentBuilder>(subDoc);
    builder->Write(u"Hello world!");
    
    subDoc->Save(get_ArtifactsDir() + u"NodeImporter.InsertAtMergeField.docx");
    
    doc->get_MailMerge()->set_FieldMergingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExNodeImporter::InsertDocumentAtMailMergeHandler>());
    
    // The main document has a merge field in it called "Document_1".
    // Execute a mail merge using a data source that contains a local system filename
    // of the document that we wish to insert into the MERGEFIELD.
    doc->get_MailMerge()->Execute(System::MakeArray<System::String>({u"Document_1"}), System::MakeArray<System::SharedPtr<System::Object>>({System::ExplicitCast<System::Object>(get_ArtifactsDir() + u"NodeImporter.InsertAtMergeField.docx")}));
    
    ASSERT_EQ(System::String(u"A document will appear here: \r") + u"Hello world!", doc->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExNodeImporter, InsertAtMergeField)
{
    s_instance->InsertAtMergeField();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
