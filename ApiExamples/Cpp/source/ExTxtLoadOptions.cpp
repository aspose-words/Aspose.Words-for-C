// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTxtLoadOptions.h"

#include <testing/test_predicates.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/io/memory_stream.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Document/TxtTrailingSpacesOptions.h>
#include <Aspose.Words.Cpp/Model/Document/TxtLoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/TxtLeadingSpacesOptions.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentDirection.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
namespace ApiExamples {

namespace gtest_test
{

class ExTxtLoadOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExTxtLoadOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExTxtLoadOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExTxtLoadOptions> ExTxtLoadOptions::s_instance;

} // namespace gtest_test

void ExTxtLoadOptions::DetectNumberingWithWhitespaces(bool detectNumberingWithWhitespaces)
{
    //ExStart
    //ExFor:TxtLoadOptions.DetectNumberingWithWhitespaces
    //ExSummary:Shows how lists are detected when plaintext documents are loaded.
    // Create a plaintext document in the form of a string with parts that may be interpreted as lists
    // Upon loading, the first three lists will always be detected by Aspose.Words, and List objects will be created for them after loading
    String textDoc = String(u"Full stop delimiters:\n") + u"1. First list item 1\n" + u"2. First list item 2\n" + u"3. First list item 3\n\n" + u"Right bracket delimiters:\n" + u"1) Second list item 1\n" + u"2) Second list item 2\n" + u"3) Second list item 3\n\n" + u"Bullet delimiters:\n" + u"• Third list item 1\n" + u"• Third list item 2\n" + u"• Third list item 3\n\n" + u"Whitespace delimiters:\n" + u"1 Fourth list item 1\n" + u"2 Fourth list item 2\n" + u"3 Fourth list item 3";

    // The fourth list, with whitespace inbetween the list number and list item contents,
    // will only be detected as a list if "DetectNumberingWithWhitespaces" in a LoadOptions object is set to true,
    // to avoid paragraphs that start with numbers being mistakenly detected as lists
    auto loadOptions = MakeObject<TxtLoadOptions>();
    loadOptions->set_DetectNumberingWithWhitespaces(detectNumberingWithWhitespaces);

    // Load the document while applying LoadOptions as a parameter and verify the result
    auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(textDoc)), loadOptions);

    if (detectNumberingWithWhitespaces)
    {
        ASSERT_EQ(4, doc->get_Lists()->get_Count());

        ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->LINQ_Any([](SharedPtr<Node> p) {
            return p->GetText().Contains(u"Fourth list") && (System::DynamicCast<Paragraph>(p))->get_IsListItem();
        }));
    }
    else
    {
        ASSERT_EQ(3, doc->get_Lists()->get_Count());

        ASSERT_FALSE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->LINQ_Any([](SharedPtr<Node> p) {
            return p->GetText().Contains(u"Fourth list") && (System::DynamicCast<Paragraph>(p))->get_IsListItem();
        }));
    }
    //ExEnd
}

namespace gtest_test
{

using DetectNumberingWithWhitespaces_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtLoadOptions::DetectNumberingWithWhitespaces)>::type;

struct DetectNumberingWithWhitespacesVP : public ExTxtLoadOptions, public ApiExamples::ExTxtLoadOptions, public ::testing::WithParamInterface<DetectNumberingWithWhitespaces_Args>
{
    static std::vector<DetectNumberingWithWhitespaces_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(DetectNumberingWithWhitespacesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DetectNumberingWithWhitespaces(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtLoadOptions, DetectNumberingWithWhitespacesVP, ::testing::ValuesIn(DetectNumberingWithWhitespacesVP::TestCases()));

} // namespace gtest_test

void ExTxtLoadOptions::TrailSpaces(Aspose::Words::TxtLeadingSpacesOptions txtLeadingSpacesOptions, Aspose::Words::TxtTrailingSpacesOptions txtTrailingSpacesOptions)
{
    //ExStart
    //ExFor:TxtLoadOptions.TrailingSpacesOptions
    //ExFor:TxtLoadOptions.LeadingSpacesOptions
    //ExFor:TxtTrailingSpacesOptions
    //ExFor:TxtLeadingSpacesOptions
    //ExSummary:Shows how to trim whitespace when loading plaintext documents.
    String textDoc = String(u"      Line 1 \n") + u"    Line 2   \n" + u" Line 3       ";

    auto loadOptions = MakeObject<TxtLoadOptions>();
    loadOptions->set_LeadingSpacesOptions(txtLeadingSpacesOptions);
    loadOptions->set_TrailingSpacesOptions(txtTrailingSpacesOptions);

    // Load the document while applying LoadOptions as a parameter and verify the result
    auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(textDoc)), loadOptions);

    SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

    switch (txtLeadingSpacesOptions)
    {
        case Aspose::Words::TxtLeadingSpacesOptions::ConvertToIndent:
            ASPOSE_ASSERT_EQ(37.8, paragraphs->idx_get(0)->get_ParagraphFormat()->get_FirstLineIndent());
            ASPOSE_ASSERT_EQ(25.2, paragraphs->idx_get(1)->get_ParagraphFormat()->get_FirstLineIndent());
            ASPOSE_ASSERT_EQ(6.3, paragraphs->idx_get(2)->get_ParagraphFormat()->get_FirstLineIndent());
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().StartsWith(u"Line 1"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().StartsWith(u"Line 2"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().StartsWith(u"Line 3"));
            break;

        case Aspose::Words::TxtLeadingSpacesOptions::Preserve:
        {
            ASSERT_TRUE(paragraphs->LINQ_All([](SharedPtr<Node> p) { return (System::DynamicCast<Paragraph>(p))->get_ParagraphFormat()->get_FirstLineIndent() == 0.0; }));
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().StartsWith(u"      Line 1"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().StartsWith(u"    Line 2"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().StartsWith(u" Line 3"));
            break;
        }

        case Aspose::Words::TxtLeadingSpacesOptions::Trim:
        {
            ASSERT_TRUE(paragraphs->LINQ_All([](SharedPtr<Node> p) { return (System::DynamicCast<Paragraph>(p))->get_ParagraphFormat()->get_FirstLineIndent() == 0.0; }));
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().StartsWith(u"Line 1"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().StartsWith(u"Line 2"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().StartsWith(u"Line 3"));
            break;
        }

    }

    switch (txtTrailingSpacesOptions)
    {
        case Aspose::Words::TxtTrailingSpacesOptions::Preserve:
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().EndsWith(u"Line 1 \r"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().EndsWith(u"Line 2   \r"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().EndsWith(u"Line 3       \f"));
            break;

        case Aspose::Words::TxtTrailingSpacesOptions::Trim:
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().EndsWith(u"Line 1\r"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().EndsWith(u"Line 2\r"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().EndsWith(u"Line 3\f"));
            break;

    }
    //ExEnd
}

namespace gtest_test
{

using TrailSpaces_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtLoadOptions::TrailSpaces)>::type;

struct TrailSpacesVP : public ExTxtLoadOptions, public ApiExamples::ExTxtLoadOptions, public ::testing::WithParamInterface<TrailSpaces_Args>
{
    static std::vector<TrailSpaces_Args> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::TxtLeadingSpacesOptions::Preserve, Aspose::Words::TxtTrailingSpacesOptions::Preserve),
            std::make_tuple(Aspose::Words::TxtLeadingSpacesOptions::ConvertToIndent, Aspose::Words::TxtTrailingSpacesOptions::Preserve),
            std::make_tuple(Aspose::Words::TxtLeadingSpacesOptions::Trim, Aspose::Words::TxtTrailingSpacesOptions::Trim),
        };
    }
};

TEST_P(TrailSpacesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TrailSpaces(get<0>(params), get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtLoadOptions, TrailSpacesVP, ::testing::ValuesIn(TrailSpacesVP::TestCases()));

} // namespace gtest_test

void ExTxtLoadOptions::DetectDocumentDirection()
{
    //ExStart
    //ExFor:TxtLoadOptions.DocumentDirection
    //ExSummary:Shows how to detect document direction automatically.
    // Create a LoadOptions object and configure it to detect text direction automatically upon loading
    auto loadOptions = MakeObject<TxtLoadOptions>();
    loadOptions->set_DocumentDirection(Aspose::Words::DocumentDirection::Auto);

    // Text like Hebrew/Arabic will be automatically detected as RTL
    auto doc = MakeObject<Document>(MyDir + u"Hebrew text.txt", loadOptions);

    ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Bidi());

    doc = MakeObject<Document>(MyDir + u"English text.txt", loadOptions);

    ASSERT_FALSE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Bidi());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTxtLoadOptions, DetectDocumentDirection)
{
    s_instance->DetectDocumentDirection();
}

} // namespace gtest_test

} // namespace ApiExamples
