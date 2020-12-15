#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlock.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockBehavior.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockCollection.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockGallery.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockType.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/GlossaryDocument.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <system/array.h>
#include <system/collections/dictionary.h>
#include <system/guid.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/string_builder.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::BuildingBlocks;

namespace ApiExamples {

class ExBuildingBlocks : public ApiExampleBase
{
public:
    //ExStart
    //ExFor:Document.GlossaryDocument
    //ExFor:BuildingBlocks.BuildingBlock
    //ExFor:BuildingBlocks.BuildingBlock.#ctor(BuildingBlocks.GlossaryDocument)
    //ExFor:BuildingBlocks.BuildingBlock.Accept(DocumentVisitor)
    //ExFor:BuildingBlocks.BuildingBlock.Behavior
    //ExFor:BuildingBlocks.BuildingBlock.Category
    //ExFor:BuildingBlocks.BuildingBlock.Description
    //ExFor:BuildingBlocks.BuildingBlock.FirstSection
    //ExFor:BuildingBlocks.BuildingBlock.Gallery
    //ExFor:BuildingBlocks.BuildingBlock.Guid
    //ExFor:BuildingBlocks.BuildingBlock.LastSection
    //ExFor:BuildingBlocks.BuildingBlock.Name
    //ExFor:BuildingBlocks.BuildingBlock.Sections
    //ExFor:BuildingBlocks.BuildingBlock.Type
    //ExFor:BuildingBlocks.BuildingBlockBehavior
    //ExFor:BuildingBlocks.BuildingBlockType
    //ExSummary:Shows how to add a custom building block to a document.
    void CreateAndInsert()
    {
        // A document's glossary document stores building blocks.
        auto doc = MakeObject<Document>();
        auto glossaryDoc = MakeObject<GlossaryDocument>();
        doc->set_GlossaryDocument(glossaryDoc);

        // Create a building block, name it, and then add it to the glossary document.
        auto block = MakeObject<BuildingBlock>(glossaryDoc);
        block->set_Name(u"Custom Block");

        glossaryDoc->AppendChild(block);

        // All new building block GUIDs have the same zero value by default, and we can give them a new unique value.
        ASSERT_EQ(u"00000000-0000-0000-0000-000000000000", System::ObjectExt::ToString(block->get_Guid()));

        block->set_Guid(System::Guid::NewGuid());

        // The following attributes categorize building blocks
        // in the menu found via Insert -> Quick Parts -> Building Blocks Organizer in Microsoft Word.
        ASSERT_EQ(u"(Empty Category)", block->get_Category());
        ASSERT_EQ(BuildingBlockType::None, block->get_Type());
        ASSERT_EQ(BuildingBlockGallery::All, block->get_Gallery());
        ASSERT_EQ(BuildingBlockBehavior::Content, block->get_Behavior());

        // Before we can add this building block to our document, we will need to give it some contents.
        // We will do that and set a category, gallery, and behavior with a document visitor.
        auto visitor = MakeObject<ExBuildingBlocks::BuildingBlockVisitor>(glossaryDoc);
        block->Accept(visitor);

        // We can access the block that we just made from the glossary document.
        SharedPtr<BuildingBlock> customBlock =
            glossaryDoc->GetBuildingBlock(BuildingBlockGallery::QuickParts, u"My custom building blocks", u"Custom Block");

        // The block itself is a section that contains the text.
        ASSERT_EQ(String::Format(u"Text inside {0}\f", customBlock->get_Name()),
                  customBlock->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText());
        ASPOSE_ASSERT_EQ(customBlock->get_FirstSection(), customBlock->get_LastSection());

        std::function<void()> parseGuid = [&customBlock]()
        {
            System::Guid::Parse(System::ObjectExt::ToString(customBlock->get_Guid()));
        };

        ASSERT_NO_THROW(parseGuid());
        //ExSkip
        ASSERT_EQ(u"My custom building blocks", customBlock->get_Category());
        //ExSkip
        ASSERT_EQ(BuildingBlockType::None, customBlock->get_Type());
        //ExSkip
        ASSERT_EQ(BuildingBlockGallery::QuickParts, customBlock->get_Gallery());
        //ExSkip
        ASSERT_EQ(BuildingBlockBehavior::Paragraph, customBlock->get_Behavior());
        //ExSkip

        // Now, we can insert it into the document as a new section.
        doc->AppendChild(doc->ImportNode(customBlock->get_FirstSection(), true));

        // We can also find it in Microsoft Word's Building Blocks Organizer and place it manually.
        doc->Save(ArtifactsDir + u"BuildingBlocks.CreateAndInsert.dotx");
    }

    /// <summary>
    /// Sets up a visited building block to be inserted into the document as a quick part and adds text to its contents.
    /// </summary>
    class BuildingBlockVisitor : public DocumentVisitor
    {
    public:
        BuildingBlockVisitor(SharedPtr<GlossaryDocument> ownerGlossaryDoc)
        {
            mBuilder = MakeObject<System::Text::StringBuilder>();
            mGlossaryDoc = ownerGlossaryDoc;
        }

        VisitorAction VisitBuildingBlockStart(SharedPtr<BuildingBlock> block) override
        {
            // Configure the building block as a quick part, and add attributes used by Building Blocks Organizer.
            block->set_Behavior(BuildingBlockBehavior::Paragraph);
            block->set_Category(u"My custom building blocks");
            block->set_Description(u"Using this block in the Quick Parts section of word will place its contents at the cursor.");
            block->set_Gallery(BuildingBlockGallery::QuickParts);

            // Add a section with text.
            // Inserting the block into the document will append this section with its child nodes at the location.
            auto section = MakeObject<Section>(mGlossaryDoc);
            block->AppendChild(section);
            block->get_FirstSection()->EnsureMinimum();

            auto run = MakeObject<Run>(mGlossaryDoc, String(u"Text inside ") + block->get_Name());
            block->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(run);

            return VisitorAction::Continue;
        }

        VisitorAction VisitBuildingBlockEnd(SharedPtr<BuildingBlock> block) override
        {
            mBuilder->Append(String(u"Visited ") + block->get_Name() + u"\r\n");
            return VisitorAction::Continue;
        }

    private:
        SharedPtr<System::Text::StringBuilder> mBuilder;
        SharedPtr<GlossaryDocument> mGlossaryDoc;
    };
    //ExEnd

    //ExStart
    //ExFor:BuildingBlocks.GlossaryDocument
    //ExFor:BuildingBlocks.GlossaryDocument.Accept(DocumentVisitor)
    //ExFor:BuildingBlocks.GlossaryDocument.BuildingBlocks
    //ExFor:BuildingBlocks.GlossaryDocument.FirstBuildingBlock
    //ExFor:BuildingBlocks.GlossaryDocument.GetBuildingBlock(BuildingBlocks.BuildingBlockGallery,System.String,System.String)
    //ExFor:BuildingBlocks.GlossaryDocument.LastBuildingBlock
    //ExFor:BuildingBlocks.BuildingBlockCollection
    //ExFor:BuildingBlocks.BuildingBlockCollection.Item(System.Int32)
    //ExFor:BuildingBlocks.BuildingBlockCollection.ToArray
    //ExFor:BuildingBlocks.BuildingBlockGallery
    //ExFor:DocumentVisitor.VisitBuildingBlockEnd(BuildingBlock)
    //ExFor:DocumentVisitor.VisitBuildingBlockStart(BuildingBlock)
    //ExFor:DocumentVisitor.VisitGlossaryDocumentEnd(GlossaryDocument)
    //ExFor:DocumentVisitor.VisitGlossaryDocumentStart(GlossaryDocument)
    //ExSummary:Shows ways of accessing building blocks in a glossary document.
    void GlossaryDocument_()
    {
        auto doc = MakeObject<Document>();
        auto glossaryDoc = MakeObject<GlossaryDocument>();

        glossaryDoc->AppendChild(
            [&]
            {
                auto tmp_0 = MakeObject<BuildingBlock>(glossaryDoc);
                tmp_0->set_Name(u"Block 1");
                return tmp_0;
            }());
        glossaryDoc->AppendChild(
            [&]
            {
                auto tmp_1 = MakeObject<BuildingBlock>(glossaryDoc);
                tmp_1->set_Name(u"Block 2");
                return tmp_1;
            }());
        glossaryDoc->AppendChild(
            [&]
            {
                auto tmp_2 = MakeObject<BuildingBlock>(glossaryDoc);
                tmp_2->set_Name(u"Block 3");
                return tmp_2;
            }());
        glossaryDoc->AppendChild(
            [&]
            {
                auto tmp_3 = MakeObject<BuildingBlock>(glossaryDoc);
                tmp_3->set_Name(u"Block 4");
                return tmp_3;
            }());
        glossaryDoc->AppendChild(
            [&]
            {
                auto tmp_4 = MakeObject<BuildingBlock>(glossaryDoc);
                tmp_4->set_Name(u"Block 5");
                return tmp_4;
            }());

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
        ASSERT_EQ(u"Block 4", glossaryDoc->GetBuildingBlock(BuildingBlockGallery::All, u"(Empty Category)", u"Block 4")->get_Name());

        // We will do that using a custom visitor,
        // which will give every BuildingBlock in the GlossaryDocument a unique GUID
        auto visitor = MakeObject<ExBuildingBlocks::GlossaryDocVisitor>();
        glossaryDoc->Accept(visitor);
        ASSERT_EQ(5, visitor->GetDictionary()->get_Count());
        //ExSkip

        std::cout << visitor->GetText() << std::endl;

        // When we open this document using Microsoft Word,
        // we can find the building blocks via Insert -> Quick Parts -> Building Blocks Organizer.
        doc->Save(ArtifactsDir + u"BuildingBlocks.GlossaryDocument.dotx");
    }

    /// <summary>
    /// Gives each building block in a visited glossary document a unique GUID, and stores the GUID-building block pairs in a dictionary.
    /// </summary>
    class GlossaryDocVisitor : public DocumentVisitor
    {
    public:
        GlossaryDocVisitor()
        {
            mBlocksByGuid = MakeObject<System::Collections::Generic::Dictionary<System::Guid, SharedPtr<BuildingBlock>>>();
            mBuilder = MakeObject<System::Text::StringBuilder>();
        }

        String GetText()
        {
            return mBuilder->ToString();
        }

        SharedPtr<System::Collections::Generic::Dictionary<System::Guid, SharedPtr<BuildingBlock>>> GetDictionary()
        {
            return mBlocksByGuid;
        }

        VisitorAction VisitGlossaryDocumentStart(SharedPtr<GlossaryDocument> glossary) override
        {
            mBuilder->AppendLine(u"Glossary document found!");
            return VisitorAction::Continue;
        }

        VisitorAction VisitGlossaryDocumentEnd(SharedPtr<GlossaryDocument> glossary) override
        {
            mBuilder->AppendLine(u"Reached end of glossary!");
            mBuilder->AppendLine(String(u"BuildingBlocks found: ") + mBlocksByGuid->get_Count());
            return VisitorAction::Continue;
        }

        VisitorAction VisitBuildingBlockStart(SharedPtr<BuildingBlock> block) override
        {
            EXPECT_EQ(u"00000000-0000-0000-0000-000000000000", System::ObjectExt::ToString(block->get_Guid()));
            //ExSkip
            block->set_Guid(System::Guid::NewGuid());
            mBlocksByGuid->Add(block->get_Guid(), block);
            return VisitorAction::Continue;
        }

        VisitorAction VisitBuildingBlockEnd(SharedPtr<BuildingBlock> block) override
        {
            mBuilder->AppendLine(String(u"\tVisited block \"") + block->get_Name() + u"\"");
            mBuilder->AppendLine(String(u"\t Type: ") + System::ObjectExt::ToString(block->get_Type()));
            mBuilder->AppendLine(String(u"\t Gallery: ") + System::ObjectExt::ToString(block->get_Gallery()));
            mBuilder->AppendLine(String(u"\t Behavior: ") + System::ObjectExt::ToString(block->get_Behavior()));
            mBuilder->AppendLine(String(u"\t Description: ") + block->get_Description());

            return VisitorAction::Continue;
        }

    private:
        SharedPtr<System::Collections::Generic::Dictionary<System::Guid, SharedPtr<BuildingBlock>>> mBlocksByGuid;
        SharedPtr<System::Text::StringBuilder> mBuilder;
    };
    //ExEnd
};

} // namespace ApiExamples
