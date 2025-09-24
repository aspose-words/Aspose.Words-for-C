// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTabStop.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/enumerator_adapter.h>
#include <system/collections/ienumerable.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/TabStopCollection.h>
#include <Aspose.Words.Cpp/Model/Text/TabStop.h>
#include <Aspose.Words.Cpp/Model/Text/TabLeader.h>
#include <Aspose.Words.Cpp/Model/Text/TabAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>

#include "TestUtil.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3952388423u, ::Aspose::Words::ApiExamples::ExTabStop, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExTabStop : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExTabStop> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExTabStop>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExTabStop> ExTabStop::s_instance;

} // namespace gtest_test

void ExTabStop::AddTabStops()
{
    //ExStart
    //ExFor:TabStopCollection.Add(TabStop)
    //ExFor:TabStopCollection.Add(Double, TabAlignment, TabLeader)
    //ExSummary:Shows how to add custom tab stops to a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto paragraph = System::ExplicitCast<Aspose::Words::Paragraph>(doc->GetChild(Aspose::Words::NodeType::Paragraph, 0, true));
    
    // Below are two ways of adding tab stops to a paragraph's collection of tab stops via the "ParagraphFormat" property.
    // 1 -  Create a "TabStop" object, and then add it to the collection:
    auto tabStop = System::MakeObject<Aspose::Words::TabStop>(Aspose::Words::ConvertUtil::InchToPoint(3), Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dashes);
    paragraph->get_ParagraphFormat()->get_TabStops()->Add(tabStop);
    
    // 2 -  Pass the values for properties of a new tab stop to the "Add" method:
    paragraph->get_ParagraphFormat()->get_TabStops()->Add(Aspose::Words::ConvertUtil::MillimeterToPoint(100), Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dashes);
    
    // Add tab stops at 5 cm to all paragraphs.
    for (auto&& para : System::IterateOver(doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Paragraph> >()))
    {
        para->get_ParagraphFormat()->get_TabStops()->Add(Aspose::Words::ConvertUtil::MillimeterToPoint(50), Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dashes);
    }
    
    // Every "tab" character takes the builder's cursor to the location of the next tab stop.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Start\tTab 1\tTab 2\tTab 3\tTab 4");
    
    doc->Save(get_ArtifactsDir() + u"TabStopCollection.AddTabStops.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"TabStopCollection.AddTabStops.docx");
    System::SharedPtr<Aspose::Words::TabStopCollection> tabStops = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_TabStops();
    
    Aspose::Words::ApiExamples::TestUtil::VerifyTabStop(141.75, Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dashes, false, tabStops->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyTabStop(216.0, Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dashes, false, tabStops->idx_get(1));
    Aspose::Words::ApiExamples::TestUtil::VerifyTabStop(283.45, Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dashes, false, tabStops->idx_get(2));
}

namespace gtest_test
{

TEST_F(ExTabStop, AddTabStops)
{
    s_instance->AddTabStops();
}

} // namespace gtest_test

void ExTabStop::TabStopCollection()
{
    //ExStart
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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::TabStopCollection> tabStops = builder->get_ParagraphFormat()->get_TabStops();
    
    // 72 points is one "inch" on the Microsoft Word tab stop ruler.
    tabStops->Add(System::MakeObject<Aspose::Words::TabStop>(72.0));
    tabStops->Add(System::MakeObject<Aspose::Words::TabStop>(432.0, Aspose::Words::TabAlignment::Right, Aspose::Words::TabLeader::Dashes));
    
    ASSERT_EQ(2, tabStops->get_Count());
    ASSERT_FALSE(tabStops->idx_get(0)->get_IsClear());
    ASSERT_FALSE(System::ObjectExt::Equals(tabStops->idx_get(0), tabStops->idx_get(1)));
    
    // Every "tab" character takes the builder's cursor to the location of the next tab stop.
    builder->Writeln(u"Start\tTab 1\tTab 2");
    
    System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    
    ASSERT_EQ(2, paragraphs->get_Count());
    
    // Each paragraph gets its tab stop collection, which clones its values from the document builder's tab stop collection.
    ASPOSE_ASSERT_EQ(paragraphs->idx_get(0)->get_ParagraphFormat()->get_TabStops(), paragraphs->idx_get(1)->get_ParagraphFormat()->get_TabStops());
    ASPOSE_ASSERT_NS(paragraphs->idx_get(0)->get_ParagraphFormat()->get_TabStops(), paragraphs->idx_get(1)->get_ParagraphFormat()->get_TabStops());
    
    // A tab stop collection can point us to TabStops before and after certain positions.
    ASPOSE_ASSERT_EQ(72.0, tabStops->Before(100.0)->get_Position());
    ASPOSE_ASSERT_EQ(432.0, tabStops->After(100.0)->get_Position());
    
    // We can clear a paragraph's tab stop collection to revert to the default tabbing behavior.
    paragraphs->idx_get(1)->get_ParagraphFormat()->get_TabStops()->Clear();
    
    ASSERT_EQ(0, paragraphs->idx_get(1)->get_ParagraphFormat()->get_TabStops()->get_Count());
    
    doc->Save(get_ArtifactsDir() + u"TabStopCollection.TabStopCollection.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"TabStopCollection.TabStopCollection.docx");
    tabStops = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_TabStops();
    
    ASSERT_EQ(2, tabStops->get_Count());
    Aspose::Words::ApiExamples::TestUtil::VerifyTabStop(72.0, Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::None, false, tabStops->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyTabStop(432.0, Aspose::Words::TabAlignment::Right, Aspose::Words::TabLeader::Dashes, false, tabStops->idx_get(1));
    
    tabStops = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_ParagraphFormat()->get_TabStops();
    
    ASSERT_EQ(0, tabStops->get_Count());
}

namespace gtest_test
{

TEST_F(ExTabStop, TabStopCollection)
{
    s_instance->TabStopCollection();
}

} // namespace gtest_test

void ExTabStop::RemoveByIndex()
{
    //ExStart
    //ExFor:TabStopCollection.RemoveByIndex
    //ExSummary:Shows how to select a tab stop in a document by its index and remove it.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::TabStopCollection> tabStops = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_TabStops();
    
    tabStops->Add(Aspose::Words::ConvertUtil::MillimeterToPoint(30), Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dashes);
    tabStops->Add(Aspose::Words::ConvertUtil::MillimeterToPoint(60), Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dashes);
    
    ASSERT_EQ(2, tabStops->get_Count());
    
    // Remove the first tab stop.
    tabStops->RemoveByIndex(0);
    
    ASSERT_EQ(1, tabStops->get_Count());
    
    doc->Save(get_ArtifactsDir() + u"TabStopCollection.RemoveByIndex.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"TabStopCollection.RemoveByIndex.docx");
    
    Aspose::Words::ApiExamples::TestUtil::VerifyTabStop(170.1, Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dashes, false, doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_TabStops()->idx_get(0));
}

namespace gtest_test
{

TEST_F(ExTabStop, RemoveByIndex)
{
    s_instance->RemoveByIndex();
}

} // namespace gtest_test

void ExTabStop::GetPositionByIndex()
{
    //ExStart
    //ExFor:TabStopCollection.GetPositionByIndex
    //ExSummary:Shows how to find a tab, stop by its index and verify its position.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::TabStopCollection> tabStops = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_TabStops();
    
    tabStops->Add(Aspose::Words::ConvertUtil::MillimeterToPoint(30), Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dashes);
    tabStops->Add(Aspose::Words::ConvertUtil::MillimeterToPoint(60), Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dashes);
    
    // Verify the position of the second tab stop in the collection.
    ASSERT_NEAR(Aspose::Words::ConvertUtil::MillimeterToPoint(60), tabStops->GetPositionByIndex(1), 0.1);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTabStop, GetPositionByIndex)
{
    s_instance->GetPositionByIndex();
}

} // namespace gtest_test

void ExTabStop::GetIndexByPosition()
{
    //ExStart
    //ExFor:TabStopCollection.GetIndexByPosition
    //ExSummary:Shows how to look up a position to see if a tab stop exists there and obtain its index.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::TabStopCollection> tabStops = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_TabStops();
    
    // Add a tab stop at a position of 30mm.
    tabStops->Add(Aspose::Words::ConvertUtil::MillimeterToPoint(30), Aspose::Words::TabAlignment::Left, Aspose::Words::TabLeader::Dashes);
    
    // A result of "0" returned by "GetIndexByPosition" confirms that a tab stop
    // at 30mm exists in this collection, and it is at index 0.
    ASSERT_EQ(0, tabStops->GetIndexByPosition(Aspose::Words::ConvertUtil::MillimeterToPoint(30)));
    
    // A "-1" returned by "GetIndexByPosition" confirms that
    // there is no tab stop in this collection with a position of 60mm.
    ASSERT_EQ(-1, tabStops->GetIndexByPosition(Aspose::Words::ConvertUtil::MillimeterToPoint(60)));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTabStop, GetIndexByPosition)
{
    s_instance->GetIndexByPosition();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
