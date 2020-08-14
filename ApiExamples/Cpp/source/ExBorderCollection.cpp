// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBorderCollection.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/collections/ienumerator.h>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
namespace ApiExamples {

namespace gtest_test
{

class ExBorderCollection : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExBorderCollection> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExBorderCollection>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExBorderCollection> ExBorderCollection::s_instance;

} // namespace gtest_test

void ExBorderCollection::GetBordersEnumerator()
{
    //ExStart
    //ExFor:BorderCollection.GetEnumerator
    //ExSummary:Shows how to enumerate all borders in a collection.
    auto doc = MakeObject<Document>(MyDir + u"Borders.docx");
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<BorderCollection> borders = builder->get_ParagraphFormat()->get_Borders();

    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Border>>> enumerator = borders->GetEnumerator();
        while (enumerator->MoveNext())
        {
            // Do something useful.
            SharedPtr<Border> b = enumerator->get_Current();
            b->set_Color(System::Drawing::Color::get_RoyalBlue());
            b->set_LineStyle(Aspose::Words::LineStyle::Double);
        }
    }

    doc->Save(ArtifactsDir + u"BorderCollection.GetBordersEnumerator.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"BorderCollection.GetBordersEnumerator.docx");

    for (auto border : System::IterateOver(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders()))
    {
        ASSERT_EQ(System::Drawing::Color::get_RoyalBlue().ToArgb(), border->get_Color().ToArgb());
        ASSERT_EQ(Aspose::Words::LineStyle::Double, border->get_LineStyle());
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
    //ExSummary:Shows how to remove all borders from a paragraph at once.
    auto doc = MakeObject<Document>(MyDir + u"Borders.docx");
    auto builder = MakeObject<DocumentBuilder>(doc);
    SharedPtr<BorderCollection> borders = builder->get_ParagraphFormat()->get_Borders();

    borders->ClearFormatting();

    doc->Save(ArtifactsDir + u"BorderCollection.RemoveAllBorders.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"BorderCollection.RemoveAllBorders.docx");

    for (auto border : System::IterateOver(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders()))
    {
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), border->get_Color().ToArgb());
        ASSERT_EQ(Aspose::Words::LineStyle::None, border->get_LineStyle());
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
