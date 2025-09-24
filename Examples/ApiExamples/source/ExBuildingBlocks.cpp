// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBuildingBlocks.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockType.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockGallery.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockCollection.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockBehavior.h>


using namespace Aspose::Words::BuildingBlocks;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2653429343u, ::Aspose::Words::ApiExamples::ExBuildingBlocks::BuildingBlockVisitor, ThisTypeBaseTypesInfo);

ExBuildingBlocks::BuildingBlockVisitor::BuildingBlockVisitor(System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> ownerGlossaryDoc)
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
    mGlossaryDoc = ownerGlossaryDoc;
}

Aspose::Words::VisitorAction ExBuildingBlocks::BuildingBlockVisitor::VisitBuildingBlockStart(System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block)
{
    // Configure the building block as a quick part, and add properties used by Building Blocks Organizer.
    block->set_Behavior(Aspose::Words::BuildingBlocks::BuildingBlockBehavior::Paragraph);
    block->set_Category(u"My custom building blocks");
    block->set_Description(u"Using this block in the Quick Parts section of word will place its contents at the cursor.");
    block->set_Gallery(Aspose::Words::BuildingBlocks::BuildingBlockGallery::QuickParts);
    
    // Add a section with text.
    // Inserting the block into the document will append this section with its child nodes at the location.
    auto section = System::MakeObject<Aspose::Words::Section>(mGlossaryDoc);
    block->AppendChild<System::SharedPtr<Aspose::Words::Section>>(section);
    block->get_FirstSection()->EnsureMinimum();
    
    auto run = System::MakeObject<Aspose::Words::Run>(mGlossaryDoc, System::String(u"Text inside ") + block->get_Name());
    block->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExBuildingBlocks::BuildingBlockVisitor::VisitBuildingBlockEnd(System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block)
{
    mBuilder->Append(System::String(u"Visited ") + block->get_Name() + u"\r\n");
    return Aspose::Words::VisitorAction::Continue;
}

RTTI_INFO_IMPL_HASH(2360311070u, ::Aspose::Words::ApiExamples::ExBuildingBlocks::GlossaryDocVisitor, ThisTypeBaseTypesInfo);

ExBuildingBlocks::GlossaryDocVisitor::GlossaryDocVisitor()
{
    mBlocksByGuid = System::MakeObject<System::Collections::Generic::Dictionary<System::Guid, System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>>();
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
}

System::String ExBuildingBlocks::GlossaryDocVisitor::GetText()
{
    return mBuilder->ToString();
}

System::SharedPtr<System::Collections::Generic::Dictionary<System::Guid, System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>> ExBuildingBlocks::GlossaryDocVisitor::GetDictionary()
{
    return mBlocksByGuid;
}

Aspose::Words::VisitorAction ExBuildingBlocks::GlossaryDocVisitor::VisitGlossaryDocumentStart(System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> glossary)
{
    mBuilder->AppendLine(u"Glossary document found!");
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExBuildingBlocks::GlossaryDocVisitor::VisitGlossaryDocumentEnd(System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> glossary)
{
    mBuilder->AppendLine(u"Reached end of glossary!");
    mBuilder->AppendLine(System::String(u"BuildingBlocks found: ") + mBlocksByGuid->get_Count());
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExBuildingBlocks::GlossaryDocVisitor::VisitBuildingBlockStart(System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block)
{
    EXPECT_EQ(u"00000000-0000-0000-0000-000000000000", System::ObjectExt::ToString(block->get_Guid()));
    //ExSkip
    block->set_Guid(System::Guid::NewGuid());
    mBlocksByGuid->Add(block->get_Guid(), block);
    return Aspose::Words::VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExBuildingBlocks::GlossaryDocVisitor::VisitBuildingBlockEnd(System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> block)
{
    mBuilder->AppendLine(System::String(u"\tVisited block \"") + block->get_Name() + u"\"");
    mBuilder->AppendLine(System::String(u"\t Type: ") + System::ObjectExt::ToString(block->get_Type()));
    mBuilder->AppendLine(System::String(u"\t Gallery: ") + System::ObjectExt::ToString(block->get_Gallery()));
    mBuilder->AppendLine(System::String(u"\t Behavior: ") + System::ObjectExt::ToString(block->get_Behavior()));
    mBuilder->AppendLine(System::String(u"\t Description: ") + block->get_Description());
    
    return Aspose::Words::VisitorAction::Continue;
}


RTTI_INFO_IMPL_HASH(3890020106u, ::Aspose::Words::ApiExamples::ExBuildingBlocks, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExBuildingBlocks : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExBuildingBlocks> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExBuildingBlocks>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExBuildingBlocks> ExBuildingBlocks::s_instance;

} // namespace gtest_test

void ExBuildingBlocks::CreateAndInsert()
{
    // A document's glossary document stores building blocks.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto glossaryDoc = System::MakeObject<Aspose::Words::BuildingBlocks::GlossaryDocument>();
    doc->set_GlossaryDocument(glossaryDoc);
    
    // Create a building block, name it, and then add it to the glossary document.
    auto block = System::MakeObject<Aspose::Words::BuildingBlocks::BuildingBlock>(glossaryDoc);
    block->set_Name(u"Custom Block");
    
    glossaryDoc->AppendChild<System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>(block);
    
    // All new building block GUIDs have the same zero value by default, and we can give them a new unique value.
    ASSERT_EQ(u"00000000-0000-0000-0000-000000000000", System::ObjectExt::ToString(block->get_Guid()));
    
    block->set_Guid(System::Guid::NewGuid());
    
    // The following properties categorize building blocks
    // in the menu we can access in Microsoft Word via "Insert" -> "Quick Parts" -> "Building Blocks Organizer".
    ASSERT_EQ(u"(Empty Category)", block->get_Category());
    ASSERT_EQ(Aspose::Words::BuildingBlocks::BuildingBlockType::None, block->get_Type());
    ASSERT_EQ(Aspose::Words::BuildingBlocks::BuildingBlockGallery::All, block->get_Gallery());
    ASSERT_EQ(Aspose::Words::BuildingBlocks::BuildingBlockBehavior::Content, block->get_Behavior());
    
    // Before we can add this building block to our document, we will need to give it some contents,
    // which we will do using a document visitor. This visitor will also set a category, gallery, and behavior.
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExBuildingBlocks::BuildingBlockVisitor>(glossaryDoc);
    // Visit start/end of the BuildingBlock.
    block->Accept(visitor);
    
    // We can access the block that we just made from the glossary document.
    System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock> customBlock = glossaryDoc->GetBuildingBlock(Aspose::Words::BuildingBlocks::BuildingBlockGallery::QuickParts, u"My custom building blocks", u"Custom Block");
    
    // The block itself is a section that contains the text.
    ASSERT_EQ(System::String::Format(u"Text inside {0}\f", customBlock->get_Name()), customBlock->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText());
    ASPOSE_ASSERT_EQ(customBlock->get_FirstSection(), customBlock->get_LastSection());
    ASSERT_NO_THROW(static_cast<std::function<void()>>([&customBlock]() -> void
    {
        System::Guid::Parse(System::ObjectExt::ToString(customBlock->get_Guid()));
    })());
    //ExSkip
    ASSERT_EQ(u"My custom building blocks", customBlock->get_Category());
    //ExSkip
    ASSERT_EQ(Aspose::Words::BuildingBlocks::BuildingBlockType::None, customBlock->get_Type());
    //ExSkip
    ASSERT_EQ(Aspose::Words::BuildingBlocks::BuildingBlockGallery::QuickParts, customBlock->get_Gallery());
    //ExSkip
    ASSERT_EQ(Aspose::Words::BuildingBlocks::BuildingBlockBehavior::Paragraph, customBlock->get_Behavior());
    //ExSkip
    
    // Now, we can insert it into the document as a new section.
    doc->AppendChild<System::SharedPtr<Aspose::Words::Node>>(doc->ImportNode(customBlock->get_FirstSection(), true));
    
    // We can also find it in Microsoft Word's Building Blocks Organizer and place it manually.
    doc->Save(get_ArtifactsDir() + u"BuildingBlocks.CreateAndInsert.dotx");
}

namespace gtest_test
{

TEST_F(ExBuildingBlocks, CreateAndInsert)
{
    s_instance->CreateAndInsert();
}

} // namespace gtest_test

void ExBuildingBlocks::GlossaryDocument()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto glossaryDoc = System::MakeObject<Aspose::Words::BuildingBlocks::GlossaryDocument>();
    
    auto child1 = System::MakeObject<Aspose::Words::BuildingBlocks::BuildingBlock>(glossaryDoc);
    child1->set_Name(u"Block 1");
    glossaryDoc->AppendChild<System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>(child1);
    auto child2 = System::MakeObject<Aspose::Words::BuildingBlocks::BuildingBlock>(glossaryDoc);
    child2->set_Name(u"Block 2");
    glossaryDoc->AppendChild<System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>(child2);
    auto child3 = System::MakeObject<Aspose::Words::BuildingBlocks::BuildingBlock>(glossaryDoc);
    child3->set_Name(u"Block 3");
    glossaryDoc->AppendChild<System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>(child3);
    auto child4 = System::MakeObject<Aspose::Words::BuildingBlocks::BuildingBlock>(glossaryDoc);
    child4->set_Name(u"Block 4");
    glossaryDoc->AppendChild<System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>(child4);
    auto child5 = System::MakeObject<Aspose::Words::BuildingBlocks::BuildingBlock>(glossaryDoc);
    child5->set_Name(u"Block 5");
    glossaryDoc->AppendChild<System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>(child5);
    
    ASSERT_EQ(5, glossaryDoc->get_BuildingBlocks()->get_Count());
    
    doc->set_GlossaryDocument(glossaryDoc);
    
    // There are various ways of accessing building blocks.
    // 1 -  Get the first/last building blocks in the collection:
    ASSERT_EQ(u"Block 1", glossaryDoc->get_FirstBuildingBlock()->get_Name());
    ASSERT_EQ(u"Block 5", glossaryDoc->get_LastBuildingBlock()->get_Name());
    
    // 2 -  Get a building block by index:
    ASSERT_EQ(u"Block 2", glossaryDoc->get_BuildingBlocks()->idx_get(1)->get_Name());
    ASSERT_EQ(u"Block 3", glossaryDoc->get_BuildingBlocks()->ToArray()->idx_get(2)->get_Name());
    
    // 3 -  Get the first building block that matches a gallery, name and category:
    ASSERT_EQ(u"Block 4", glossaryDoc->GetBuildingBlock(Aspose::Words::BuildingBlocks::BuildingBlockGallery::All, u"(Empty Category)", u"Block 4")->get_Name());
    
    // We will do that using a custom visitor,
    // which will give every BuildingBlock in the GlossaryDocument a unique GUID
    auto visitor = System::MakeObject<Aspose::Words::ApiExamples::ExBuildingBlocks::GlossaryDocVisitor>();
    // Visit start/end of the Glossary document.
    glossaryDoc->Accept(visitor);
    // Visit only start of the Glossary document.
    glossaryDoc->AcceptStart(visitor);
    // Visit only end of the Glossary document.
    glossaryDoc->AcceptEnd(visitor);
    ASSERT_EQ(5, visitor->GetDictionary()->get_Count());
    //ExSkip
    
    std::cout << visitor->GetText() << std::endl;
    
    // In Microsoft Word, we can access the building blocks via "Insert" -> "Quick Parts" -> "Building Blocks Organizer".
    doc->Save(get_ArtifactsDir() + u"BuildingBlocks.GlossaryDocument.dotx");
}

namespace gtest_test
{

TEST_F(ExBuildingBlocks, GlossaryDocument)
{
    s_instance->GlossaryDocument();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
