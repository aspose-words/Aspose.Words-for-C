// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBuildingBlocks.h"

#include <testing/test_predicates.h>
#include <system/text/string_builder.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/guid.h>
#include <system/console.h>
#include <system/collections/dictionary.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/GlossaryDocument.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockType.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockGallery.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockCollection.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockBehavior.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlock.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::BuildingBlocks;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(359510756u, ::ApiExamples::ExBuildingBlocks::GlossaryDocVisitor, ThisTypeBaseTypesInfo);

ExBuildingBlocks::GlossaryDocVisitor::GlossaryDocVisitor()
{
    mBlocksByGuid = MakeObject<System::Collections::Generic::Dictionary<System::Guid, SharedPtr<BuildingBlock>>>();
    mBuilder = MakeObject<System::Text::StringBuilder>();
}

String ExBuildingBlocks::GlossaryDocVisitor::GetText()
{
    return mBuilder->ToString();
}

SharedPtr<System::Collections::Generic::Dictionary<System::Guid, SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>> ExBuildingBlocks::GlossaryDocVisitor::GetDictionary()
{
    return mBlocksByGuid;
}

Aspose::Words::VisitorAction ExBuildingBlocks::GlossaryDocVisitor::VisitGlossaryDocumentStart(SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> glossary)
{
    mBuilder->AppendLine(u"Glossary document found!");
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExBuildingBlocks::GlossaryDocVisitor::VisitGlossaryDocumentEnd(SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> glossary)
{
    mBuilder->AppendLine(u"Reached end of glossary!");
    mBuilder->AppendLine(String(u"BuildingBlocks found: ") + mBlocksByGuid->get_Count());
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExBuildingBlocks::GlossaryDocVisitor::VisitBuildingBlockStart(SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block)
{
    EXPECT_EQ(u"00000000-0000-0000-0000-000000000000", System::ObjectExt::ToString(block->get_Guid()));
    //ExSkip
    block->set_Guid(System::Guid::NewGuid());
    mBlocksByGuid->Add(block->get_Guid(), block);
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExBuildingBlocks::GlossaryDocVisitor::VisitBuildingBlockEnd(SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block)
{
    mBuilder->AppendLine(String(u"\tVisited block \"") + block->get_Name() + u"\"");
    mBuilder->AppendLine(String(u"\t Type: ") + System::ObjectExt::ToString(block->get_Type()));
    mBuilder->AppendLine(String(u"\t Gallery: ") + System::ObjectExt::ToString(block->get_Gallery()));
    mBuilder->AppendLine(String(u"\t Behavior: ") + System::ObjectExt::ToString(block->get_Behavior()));
    mBuilder->AppendLine(String(u"\t Description: ") + block->get_Description());

    return Aspose::Words::VisitorAction::Continue;
}

System::Object::shared_members_type ApiExamples::ExBuildingBlocks::GlossaryDocVisitor::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExBuildingBlocks::GlossaryDocVisitor::mBlocksByGuid", this->mBlocksByGuid);
    result.Add("ApiExamples::ExBuildingBlocks::GlossaryDocVisitor::mBuilder", this->mBuilder);

    return result;
}

RTTI_INFO_IMPL_HASH(580636965u, ::ApiExamples::ExBuildingBlocks::BuildingBlockVisitor, ThisTypeBaseTypesInfo);

ExBuildingBlocks::BuildingBlockVisitor::BuildingBlockVisitor(SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> ownerGlossaryDoc)
{
    mBuilder = MakeObject<System::Text::StringBuilder>();
    mGlossaryDoc = ownerGlossaryDoc;
}

Aspose::Words::VisitorAction ExBuildingBlocks::BuildingBlockVisitor::VisitBuildingBlockStart(SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block)
{
    // Change values by default of created BuildingBlock
    block->set_Behavior(Aspose::Words::BuildingBlocks::BuildingBlockBehavior::Paragraph);
    block->set_Category(u"My custom building blocks");
    block->set_Description(u"Using this block in the Quick Parts section of word will place its contents at the cursor.");
    block->set_Gallery(Aspose::Words::BuildingBlocks::BuildingBlockGallery::QuickParts);

    block->set_Guid(System::Guid::NewGuid());

    // Add content for the BuildingBlock to have an effect when used in the document
    auto section = MakeObject<Section>(mGlossaryDoc);
    block->AppendChild(section);

    auto body = MakeObject<Body>(mGlossaryDoc);
    section->AppendChild(body);

    auto paragraph = MakeObject<Paragraph>(mGlossaryDoc);
    body->AppendChild(paragraph);

    // Add text that will be visible in the document
    auto run = MakeObject<Run>(mGlossaryDoc, String(u"Text inside ") + block->get_Name());
    block->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(run);

    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExBuildingBlocks::BuildingBlockVisitor::VisitBuildingBlockEnd(SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block)
{
    mBuilder->Append(String(u"Visited ") + block->get_Name() + u"\r\n");
    return Aspose::Words::VisitorAction::Continue;
}

System::Object::shared_members_type ApiExamples::ExBuildingBlocks::BuildingBlockVisitor::GetSharedMembers()
{
    auto result = Aspose::Words::DocumentVisitor::GetSharedMembers();

    result.Add("ApiExamples::ExBuildingBlocks::BuildingBlockVisitor::mBuilder", this->mBuilder);
    result.Add("ApiExamples::ExBuildingBlocks::BuildingBlockVisitor::mGlossaryDoc", this->mGlossaryDoc);

    return result;
}

namespace gtest_test
{

class ExBuildingBlocks : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExBuildingBlocks> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExBuildingBlocks>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExBuildingBlocks> ExBuildingBlocks::s_instance;

} // namespace gtest_test

void ExBuildingBlocks::BuildingBlockFields()
{
    auto doc = MakeObject<Document>();

    // BuildingBlocks are stored inside the glossary document
    // If you're making a document from scratch, the glossary document must also be manually created
    auto glossaryDoc = MakeObject<Aspose::Words::BuildingBlocks::GlossaryDocument>();
    doc->set_GlossaryDocument(glossaryDoc);

    // Create a building block and name it
    auto block = MakeObject<BuildingBlock>(glossaryDoc);
    block->set_Name(u"Custom Block");

    // Put in in the document's glossary document
    glossaryDoc->AppendChild(block);
    ASSERT_EQ(1, glossaryDoc->get_Count());

    // All GUIDs are this value by default
    ASSERT_EQ(u"00000000-0000-0000-0000-000000000000", System::ObjectExt::ToString(block->get_Guid()));

    // In Microsoft Word, we can use these attributes to find blocks in Insert > Quick Parts > Building Blocks Organizer
    ASSERT_EQ(u"(Empty Category)", block->get_Category());
    ASSERT_EQ(Aspose::Words::BuildingBlocks::BuildingBlockType::None, block->get_Type());
    ASSERT_EQ(Aspose::Words::BuildingBlocks::BuildingBlockGallery::All, block->get_Gallery());
    ASSERT_EQ(Aspose::Words::BuildingBlocks::BuildingBlockBehavior::Content, block->get_Behavior());

    // If we want to use our building block as an AutoText quick part, we need to give it some text and change some properties
    // All the necessary preparation will be done in a custom document visitor that we will accept
    auto visitor = MakeObject<ExBuildingBlocks::BuildingBlockVisitor>(glossaryDoc);
    block->Accept(visitor);

    // We can find the block we made in the glossary document like this
    SharedPtr<BuildingBlock> customBlock = glossaryDoc->GetBuildingBlock(Aspose::Words::BuildingBlocks::BuildingBlockGallery::QuickParts, u"My custom building blocks", u"Custom Block");

    // Our block contains one section which now contains our text
    ASSERT_EQ(String::Format(u"Text inside {0}\f",customBlock->get_Name()), customBlock->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText());
    ASPOSE_ASSERT_EQ(customBlock->get_FirstSection(), customBlock->get_LastSection());

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_0 = [&customBlock]()
    {
        System::Guid::Parse(System::ObjectExt::ToString(customBlock->get_Guid()));
    };

    ASSERT_NO_THROW(_local_func_0());
    //ExSkip
    ASSERT_EQ(u"My custom building blocks", customBlock->get_Category());
    //ExSkip
    ASSERT_EQ(Aspose::Words::BuildingBlocks::BuildingBlockType::None, customBlock->get_Type());
    //ExSkip
    ASSERT_EQ(Aspose::Words::BuildingBlocks::BuildingBlockGallery::QuickParts, customBlock->get_Gallery());
    //ExSkip
    ASSERT_EQ(Aspose::Words::BuildingBlocks::BuildingBlockBehavior::Paragraph, customBlock->get_Behavior());
    //ExSkip

    // Then we can insert it into the document as a new section
    doc->AppendChild(doc->ImportNode(customBlock->get_FirstSection(), true));

    // Or we can find it in Microsoft Word's Building Blocks Organizer and place it manually
    doc->Save(ArtifactsDir + u"BuildingBlocks.BuildingBlockFields.dotx");
}

namespace gtest_test
{

TEST_F(ExBuildingBlocks, BuildingBlockFields)
{
    s_instance->BuildingBlockFields();
}

} // namespace gtest_test

void ExBuildingBlocks::GlossaryDocument()
{
    auto doc = MakeObject<Document>();

    auto glossaryDoc = MakeObject<Aspose::Words::BuildingBlocks::GlossaryDocument>();
    glossaryDoc->AppendChild([&]{ auto tmp_0 = MakeObject<BuildingBlock>(glossaryDoc); tmp_0->set_Name(u"Block 1"); return tmp_0; }());
    glossaryDoc->AppendChild([&]{ auto tmp_1 = MakeObject<BuildingBlock>(glossaryDoc); tmp_1->set_Name(u"Block 2"); return tmp_1; }());
    glossaryDoc->AppendChild([&]{ auto tmp_2 = MakeObject<BuildingBlock>(glossaryDoc); tmp_2->set_Name(u"Block 3"); return tmp_2; }());
    glossaryDoc->AppendChild([&]{ auto tmp_3 = MakeObject<BuildingBlock>(glossaryDoc); tmp_3->set_Name(u"Block 4"); return tmp_3; }());
    glossaryDoc->AppendChild([&]{ auto tmp_4 = MakeObject<BuildingBlock>(glossaryDoc); tmp_4->set_Name(u"Block 5"); return tmp_4; }());

    ASSERT_EQ(5, glossaryDoc->get_BuildingBlocks()->get_Count());

    doc->set_GlossaryDocument(glossaryDoc);

    // There is a different ways how to get created building blocks
    ASSERT_EQ(u"Block 1", glossaryDoc->get_FirstBuildingBlock()->get_Name());
    ASSERT_EQ(u"Block 2", glossaryDoc->get_BuildingBlocks()->idx_get(1)->get_Name());
    ASSERT_EQ(u"Block 3", glossaryDoc->get_BuildingBlocks()->ToArray()->idx_get(2)->get_Name());
    ASSERT_EQ(u"Block 4", glossaryDoc->GetBuildingBlock(Aspose::Words::BuildingBlocks::BuildingBlockGallery::All, u"(Empty Category)", u"Block 4")->get_Name());
    ASSERT_EQ(u"Block 5", glossaryDoc->get_LastBuildingBlock()->get_Name());

    // We will do that using a custom visitor, which also will give every BuildingBlock in the GlossaryDocument a unique GUID
    auto visitor = MakeObject<ExBuildingBlocks::GlossaryDocVisitor>();
    glossaryDoc->Accept(visitor);
    ASSERT_EQ(5, visitor->GetDictionary()->get_Count());
    //ExSkip

    System::Console::WriteLine(visitor->GetText());

    // We can find our new blocks in Microsoft Word via Insert > Quick Parts > Building Blocks Organizer...
    doc->Save(ArtifactsDir + u"BuildingBlocks.GlossaryDocument.dotx");
}

namespace gtest_test
{

TEST_F(ExBuildingBlocks, GlossaryDocument)
{
    s_instance->GlossaryDocument();
}

} // namespace gtest_test

} // namespace ApiExamples
