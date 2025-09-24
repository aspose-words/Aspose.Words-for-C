// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBorderCollection.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/collections/ienumerator.h>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(116241466u, ::Aspose::Words::ApiExamples::ExBorderCollection, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExBorderCollection : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExBorderCollection> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExBorderCollection>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExBorderCollection> ExBorderCollection::s_instance;

} // namespace gtest_test

void ExBorderCollection::GetBordersEnumerator()
{
    //ExStart
    //ExFor:BorderCollection.GetEnumerator
    //ExSummary:Shows how to iterate over and edit all of the borders in a paragraph format object.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Configure the builder's paragraph format settings to create a green wave border on all sides.
    System::SharedPtr<Aspose::Words::BorderCollection> borders = builder->get_ParagraphFormat()->get_Borders();
    
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Border>>> enumerator = borders->GetEnumerator();
        while (enumerator->MoveNext())
        {
            System::SharedPtr<Aspose::Words::Border> border = enumerator->get_Current();
            border->set_Color(System::Drawing::Color::get_Green());
            border->set_LineStyle(Aspose::Words::LineStyle::Wave);
            border->set_LineWidth(3);
        }
    }
    
    // Insert a paragraph. Our border settings will determine the appearance of its border.
    builder->Writeln(u"Hello world!");
    
    doc->Save(get_ArtifactsDir() + u"BorderCollection.GetBordersEnumerator.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"BorderCollection.GetBordersEnumerator.docx");
    
    for (auto&& border : System::IterateOver(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders()))
    {
        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), border->get_Color().ToArgb());
        ASSERT_EQ(Aspose::Words::LineStyle::Wave, border->get_LineStyle());
        ASPOSE_ASSERT_EQ(3.0, border->get_LineWidth());
    }
}

namespace gtest_test
{

TEST_F(ExBorderCollection, GetBordersEnumerator)
{
    s_instance->GetBordersEnumerator();
}

} // namespace gtest_test

void ExBorderCollection::RemoveAllBorders()
{
    //ExStart
    //ExFor:BorderCollection.ClearFormatting
    //ExSummary:Shows how to remove all borders from all paragraphs in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Borders.docx");
    
    // The first paragraph of this document has visible borders with these settings.
    System::SharedPtr<Aspose::Words::BorderCollection> firstParagraphBorders = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders();
    
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), firstParagraphBorders->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::LineStyle::Single, firstParagraphBorders->get_LineStyle());
    ASPOSE_ASSERT_EQ(3.0, firstParagraphBorders->get_LineWidth());
    
    // Use the "ClearFormatting" method on each paragraph to remove all borders.
    for (auto&& paragraph : System::IterateOver<Aspose::Words::Paragraph>(doc->get_FirstSection()->get_Body()->get_Paragraphs()))
    {
        paragraph->get_ParagraphFormat()->get_Borders()->ClearFormatting();
        
        for (auto&& border : System::IterateOver(paragraph->get_ParagraphFormat()->get_Borders()))
        {
            ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), border->get_Color().ToArgb());
            ASSERT_EQ(Aspose::Words::LineStyle::None, border->get_LineStyle());
            ASPOSE_ASSERT_EQ(0.0, border->get_LineWidth());
        }
    }
    
    doc->Save(get_ArtifactsDir() + u"BorderCollection.RemoveAllBorders.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"BorderCollection.RemoveAllBorders.docx");
    
    for (auto&& border : System::IterateOver(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders()))
    {
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), border->get_Color().ToArgb());
        ASSERT_EQ(Aspose::Words::LineStyle::None, border->get_LineStyle());
        ASPOSE_ASSERT_EQ(0.0, border->get_LineWidth());
    }
}

namespace gtest_test
{

TEST_F(ExBorderCollection, RemoveAllBorders)
{
    s_instance->RemoveAllBorders();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
