#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/TabAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/TabLeader.h>
#include <Aspose.Words.Cpp/Model/Text/TabStop.h>
#include <Aspose.Words.Cpp/Model/Text/TabStopCollection.h>
#include <system/collections/ienumerable.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;

namespace ApiExamples {

class ExTabStop : public ApiExampleBase
{
public:
    void AddTabStops()
    {
        //ExStart
        //ExFor:TabStopCollection.Add(TabStop)
        //ExFor:TabStopCollection.Add(Double, TabAlignment, TabLeader)
        //ExSummary:Shows how to add tab stops to a document and set their positions.
        auto doc = MakeObject<Document>();
        auto paragraph = System::DynamicCast<Paragraph>(doc->GetChild(NodeType::Paragraph, 0, true));

        // Create a TabStop object and add it to the document
        auto tabStop = MakeObject<TabStop>(ConvertUtil::InchToPoint(3), TabAlignment::Left, TabLeader::Dashes);
        paragraph->get_ParagraphFormat()->get_TabStops()->Add(tabStop);

        // Add a tab stop without explicitly creating new TabStop objects
        paragraph->get_ParagraphFormat()->get_TabStops()->Add(ConvertUtil::MillimeterToPoint(100), TabAlignment::Left, TabLeader::Dashes);

        // Add tab stops at 5 cm to all paragraphs
        for (auto para : System::IterateOver(doc->GetChildNodes(NodeType::Paragraph, true)->LINQ_OfType<SharedPtr<Paragraph>>()))
        {
            para->get_ParagraphFormat()->get_TabStops()->Add(ConvertUtil::MillimeterToPoint(50), TabAlignment::Left, TabLeader::Dashes);
        }

        // Insert text with tabs that demonstrate the tab stops
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Start\tTab 1\tTab 2\tTab 3\tTab 4");

        doc->Save(ArtifactsDir + u"TabStopCollection.AddTabStops.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"TabStopCollection.AddTabStops.docx");
        SharedPtr<TabStopCollection> tabStops =
            doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_TabStops();

        TestUtil::VerifyTabStop(141.75, TabAlignment::Left, TabLeader::Dashes, false, tabStops->idx_get(0));
        TestUtil::VerifyTabStop(216.0, TabAlignment::Left, TabLeader::Dashes, false, tabStops->idx_get(1));
        TestUtil::VerifyTabStop(283.45, TabAlignment::Left, TabLeader::Dashes, false, tabStops->idx_get(2));
    }

    void TabStopCollection_()
    {
        //ExStart
        //ExFor:TabStop.#ctor
        //ExFor:TabStop.#ctor(Double)
        //ExFor:TabStop.#ctor(Double,TabAlignment,TabLeader)
        //ExFor:TabStop.Equals(TabStop)
        //ExFor:TabStop.IsClear
        //ExFor:TabStopCollection
        //ExFor:TabStopCollection.After(Double)
        //ExFor:TabStopCollection.Before(Double)
        //ExFor:TabStopCollection.Clear
        //ExFor:TabStopCollection.Count
        //ExFor:TabStopCollection.Equals(TabStopCollection)
        //ExFor:TabStopCollection.Equals(Object)
        //ExFor:TabStopCollection.GetHashCode
        //ExFor:TabStopCollection.Item(Double)
        //ExFor:TabStopCollection.Item(Int32)
        //ExSummary:Shows how to work with a document's collection of tab stops.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Access the collection of tab stops and add some tab stops to it
        SharedPtr<TabStopCollection> tabStops = builder->get_ParagraphFormat()->get_TabStops();

        // 72 points is one "inch" on the Microsoft Word tab stop ruler
        tabStops->Add(MakeObject<TabStop>(72.0));
        tabStops->Add(MakeObject<TabStop>(432.0, TabAlignment::Right, TabLeader::Dashes));

        ASSERT_EQ(2, tabStops->get_Count());
        ASSERT_FALSE(tabStops->idx_get(0)->get_IsClear());
        ASSERT_FALSE(System::ObjectExt::Equals(tabStops->idx_get(0), tabStops->idx_get(1)));

        // Every "tab" character takes the builder's cursor to the next tab stop
        builder->Writeln(u"Start\tTab 1\tTab 2");

        // Get the collection of paragraphs that we have created
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
        ASSERT_EQ(2, paragraphs->get_Count());

        // Each paragraph gets its own TabStopCollection which gets values from the DocumentBuilder's collection
        ASPOSE_ASSERT_EQ(paragraphs->idx_get(0)->get_ParagraphFormat()->get_TabStops(), paragraphs->idx_get(1)->get_ParagraphFormat()->get_TabStops());
        ASPOSE_ASSERT_NS(paragraphs->idx_get(0)->get_ParagraphFormat()->get_TabStops(), paragraphs->idx_get(1)->get_ParagraphFormat()->get_TabStops());

        // A TabStopCollection can point us to TabStops before and after certain positions
        ASPOSE_ASSERT_EQ(72.0, tabStops->Before(100.0)->get_Position());
        ASPOSE_ASSERT_EQ(432.0, tabStops->After(100.0)->get_Position());

        // We can clear a paragraph's TabStopCollection to revert to the default tabbing behaviour
        paragraphs->idx_get(1)->get_ParagraphFormat()->get_TabStops()->Clear();

        ASSERT_EQ(0, paragraphs->idx_get(1)->get_ParagraphFormat()->get_TabStops()->get_Count());

        doc->Save(ArtifactsDir + u"TabStopCollection.TabStopCollection.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"TabStopCollection.TabStopCollection.docx");
        tabStops = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_TabStops();

        ASSERT_EQ(2, tabStops->get_Count());
        TestUtil::VerifyTabStop(72.0, TabAlignment::Left, TabLeader::None, false, tabStops->idx_get(0));
        TestUtil::VerifyTabStop(432.0, TabAlignment::Right, TabLeader::Dashes, false, tabStops->idx_get(1));

        tabStops = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_ParagraphFormat()->get_TabStops();

        ASSERT_EQ(0, tabStops->get_Count());
    }

    void RemoveByIndex()
    {
        //ExStart
        //ExFor:TabStopCollection.RemoveByIndex
        //ExSummary:Shows how to select a tab stop in a document by its index and remove it.
        auto doc = MakeObject<Document>();
        SharedPtr<TabStopCollection> tabStops =
            doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_TabStops();

        tabStops->Add(ConvertUtil::MillimeterToPoint(30), TabAlignment::Left, TabLeader::Dashes);
        tabStops->Add(ConvertUtil::MillimeterToPoint(60), TabAlignment::Left, TabLeader::Dashes);

        ASSERT_EQ(2, tabStops->get_Count());

        // Tab stop placed at 30 mm is removed
        tabStops->RemoveByIndex(0);

        ASSERT_EQ(1, tabStops->get_Count());

        doc->Save(ArtifactsDir + u"TabStopCollection.RemoveByIndex.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"TabStopCollection.RemoveByIndex.docx");

        TestUtil::VerifyTabStop(170.1, TabAlignment::Left, TabLeader::Dashes, false,
                                doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_TabStops()->idx_get(0));
    }

    void GetPositionByIndex()
    {
        //ExStart
        //ExFor:TabStopCollection.GetPositionByIndex
        //ExSummary:Shows how to find a tab stop by it's index and get its position.
        auto doc = MakeObject<Document>();
        SharedPtr<TabStopCollection> tabStops =
            doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_TabStops();

        tabStops->Add(ConvertUtil::MillimeterToPoint(30), TabAlignment::Left, TabLeader::Dashes);
        tabStops->Add(ConvertUtil::MillimeterToPoint(60), TabAlignment::Left, TabLeader::Dashes);

        // Get the position of the second tab stop in the collection
        ASSERT_NEAR(ConvertUtil::MillimeterToPoint(60), tabStops->GetPositionByIndex(1), 0.1);
        //ExEnd
    }

    void GetIndexByPosition()
    {
        //ExStart
        //ExFor:TabStopCollection.GetIndexByPosition
        //ExSummary:Shows how to look up a position to see if a tab stop exists there, and if so, obtain its index.
        auto doc = MakeObject<Document>();
        SharedPtr<TabStopCollection> tabStops =
            doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_TabStops();

        // Add a tab stop at a position of 30mm
        tabStops->Add(ConvertUtil::MillimeterToPoint(30), TabAlignment::Left, TabLeader::Dashes);

        // "0" confirms that a tab stop at 30mm exists in this collection, and it is at index 0
        ASSERT_EQ(0, tabStops->GetIndexByPosition(ConvertUtil::MillimeterToPoint(30)));

        // "-1" means that there is no tab stop in this collection with a position of 60mm
        ASSERT_EQ(-1, tabStops->GetIndexByPosition(ConvertUtil::MillimeterToPoint(60)));
        //ExEnd
    }
};

} // namespace ApiExamples
