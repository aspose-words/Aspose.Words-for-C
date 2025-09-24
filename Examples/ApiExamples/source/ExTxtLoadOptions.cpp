// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTxtLoadOptions.h"

#include <testing/test_predicates.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/memory_stream.h>
#include <system/func.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Loading/TxtLoadOptions.h>
#include <Aspose.Words.Cpp/Model/Loading/DocumentDirection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Loading;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3034698072u, ::Aspose::Words::ApiExamples::ExTxtLoadOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExTxtLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExTxtLoadOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExTxtLoadOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExTxtLoadOptions> ExTxtLoadOptions::s_instance;

} // namespace gtest_test

void ExTxtLoadOptions::DetectNumberingWithWhitespaces(bool detectNumberingWithWhitespaces)
{
    //ExStart
    //ExFor:TxtLoadOptions.DetectNumberingWithWhitespaces
    //ExSummary:Shows how to detect lists when loading plaintext documents.
    // Create a plaintext document in a string with four separate parts that we may interpret as lists,
    // with different delimiters. Upon loading the plaintext document into a "Document" object,
    // Aspose.Words will always detect the first three lists and will add a "List" object
    // for each to the document's "Lists" property.
    const System::String textDoc = System::String(u"Full stop delimiters:\n") + u"1. First list item 1\n" + u"2. First list item 2\n" + u"3. First list item 3\n\n" + u"Right bracket delimiters:\n" + u"1) Second list item 1\n" + u"2) Second list item 2\n" + u"3) Second list item 3\n\n" + u"Bullet delimiters:\n" + u"• Third list item 1\n" + u"• Third list item 2\n" + u"• Third list item 3\n\n" + u"Whitespace delimiters:\n" + u"1 Fourth list item 1\n" + u"2 Fourth list item 2\n" + u"3 Fourth list item 3";
    
    // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
    // to modify how we load a plaintext document.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::TxtLoadOptions>();
    
    // Set the "DetectNumberingWithWhitespaces" property to "true" to detect numbered items
    // with whitespace delimiters, such as the fourth list in our document, as lists.
    // This may also falsely detect paragraphs that begin with numbers as lists.
    // Set the "DetectNumberingWithWhitespaces" property to "false"
    // to not create lists from numbered items with whitespace delimiters.
    loadOptions->set_DetectNumberingWithWhitespaces(detectNumberingWithWhitespaces);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(System::MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(textDoc)), loadOptions);
    
    if (detectNumberingWithWhitespaces)
    {
        ASSERT_EQ(4, doc->get_Lists()->get_Count());
        ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> p)>>([](System::SharedPtr<Aspose::Words::Node> p) -> bool
        {
            return p->GetText().Contains(u"Fourth list") && (System::ExplicitCast<Aspose::Words::Paragraph>(p))->get_IsListItem();
        }))));
    }
    else
    {
        ASSERT_EQ(3, doc->get_Lists()->get_Count());
        ASSERT_FALSE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> p)>>([](System::SharedPtr<Aspose::Words::Node> p) -> bool
        {
            return p->GetText().Contains(u"Fourth list") && (System::ExplicitCast<Aspose::Words::Paragraph>(p))->get_IsListItem();
        }))));
    }
    //ExEnd
}

namespace gtest_test
{

using ExTxtLoadOptions_DetectNumberingWithWhitespaces_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExTxtLoadOptions::DetectNumberingWithWhitespaces)>::type;

struct ExTxtLoadOptions_DetectNumberingWithWhitespaces : public ExTxtLoadOptions, public Aspose::Words::ApiExamples::ExTxtLoadOptions, public ::testing::WithParamInterface<ExTxtLoadOptions_DetectNumberingWithWhitespaces_Args>
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

TEST_P(ExTxtLoadOptions_DetectNumberingWithWhitespaces, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DetectNumberingWithWhitespaces(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtLoadOptions_DetectNumberingWithWhitespaces, ::testing::ValuesIn(ExTxtLoadOptions_DetectNumberingWithWhitespaces::TestCases()));

} // namespace gtest_test

void ExTxtLoadOptions::TrailSpaces(Aspose::Words::Loading::TxtLeadingSpacesOptions txtLeadingSpacesOptions, Aspose::Words::Loading::TxtTrailingSpacesOptions txtTrailingSpacesOptions)
{
    //ExStart
    //ExFor:TxtLoadOptions.TrailingSpacesOptions
    //ExFor:TxtLoadOptions.LeadingSpacesOptions
    //ExFor:TxtTrailingSpacesOptions
    //ExFor:TxtLeadingSpacesOptions
    //ExSummary:Shows how to trim whitespace when loading plaintext documents.
    System::String textDoc = System::String(u"      Line 1 \n") + u"    Line 2   \n" + u" Line 3       ";
    
    // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
    // to modify how we load a plaintext document.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::TxtLoadOptions>();
    
    // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.Preserve"
    // to preserve all whitespace characters at the start of every line.
    // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.ConvertToIndent"
    // to remove all whitespace characters from the start of every line,
    // and then apply a left first line indent to the paragraph to simulate the effect of the whitespaces.
    // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.Trim"
    // to remove all whitespace characters from every line's start.
    loadOptions->set_LeadingSpacesOptions(txtLeadingSpacesOptions);
    
    // Set the "TrailingSpacesOptions" property to "TxtTrailingSpacesOptions.Preserve"
    // to preserve all whitespace characters at the end of every line. 
    // Set the "TrailingSpacesOptions" property to "TxtTrailingSpacesOptions.Trim" to 
    // remove all whitespace characters from the end of every line.
    loadOptions->set_TrailingSpacesOptions(txtTrailingSpacesOptions);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(System::MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(textDoc)), loadOptions);
    System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    
    switch (txtLeadingSpacesOptions)
    {
        case Aspose::Words::Loading::TxtLeadingSpacesOptions::ConvertToIndent:
            ASPOSE_ASSERT_EQ(37.8, paragraphs->idx_get(0)->get_ParagraphFormat()->get_FirstLineIndent());
            ASPOSE_ASSERT_EQ(25.2, paragraphs->idx_get(1)->get_ParagraphFormat()->get_FirstLineIndent());
            ASPOSE_ASSERT_EQ(6.3, paragraphs->idx_get(2)->get_ParagraphFormat()->get_FirstLineIndent());
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().StartsWith(u"Line 1"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().StartsWith(u"Line 2"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().StartsWith(u"Line 3"));
            break;
        
        case Aspose::Words::Loading::TxtLeadingSpacesOptions::Preserve:
            ASSERT_TRUE(paragraphs->LINQ_All(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> p)>>([](System::SharedPtr<Aspose::Words::Node> p) -> bool
            {
                return (System::ExplicitCast<Aspose::Words::Paragraph>(p))->get_ParagraphFormat()->get_FirstLineIndent() == 0.0;
            }))));
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().StartsWith(u"      Line 1"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().StartsWith(u"    Line 2"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().StartsWith(u" Line 3"));
            break;
        
        case Aspose::Words::Loading::TxtLeadingSpacesOptions::Trim:
            ASSERT_TRUE(paragraphs->LINQ_All(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> p)>>([](System::SharedPtr<Aspose::Words::Node> p) -> bool
            {
                return (System::ExplicitCast<Aspose::Words::Paragraph>(p))->get_ParagraphFormat()->get_FirstLineIndent() == 0.0;
            }))));
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().StartsWith(u"Line 1"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().StartsWith(u"Line 2"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().StartsWith(u"Line 3"));
            break;
        
    }
    
    switch (txtTrailingSpacesOptions)
    {
        case Aspose::Words::Loading::TxtTrailingSpacesOptions::Preserve:
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().EndsWith(u"Line 1 \r"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().EndsWith(u"Line 2   \r"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().EndsWith(u"Line 3       \f"));
            break;
        
        case Aspose::Words::Loading::TxtTrailingSpacesOptions::Trim:
            ASSERT_TRUE(paragraphs->idx_get(0)->GetText().EndsWith(u"Line 1\r"));
            ASSERT_TRUE(paragraphs->idx_get(1)->GetText().EndsWith(u"Line 2\r"));
            ASSERT_TRUE(paragraphs->idx_get(2)->GetText().EndsWith(u"Line 3\f"));
            break;
        
    }
    //ExEnd
}

namespace gtest_test
{

using ExTxtLoadOptions_TrailSpaces_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExTxtLoadOptions::TrailSpaces)>::type;

struct ExTxtLoadOptions_TrailSpaces : public ExTxtLoadOptions, public Aspose::Words::ApiExamples::ExTxtLoadOptions, public ::testing::WithParamInterface<ExTxtLoadOptions_TrailSpaces_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Loading::TxtLeadingSpacesOptions::Preserve, Aspose::Words::Loading::TxtTrailingSpacesOptions::Preserve),
            std::make_tuple(Aspose::Words::Loading::TxtLeadingSpacesOptions::ConvertToIndent, Aspose::Words::Loading::TxtTrailingSpacesOptions::Preserve),
            std::make_tuple(Aspose::Words::Loading::TxtLeadingSpacesOptions::Trim, Aspose::Words::Loading::TxtTrailingSpacesOptions::Trim),
        };
    }
};

TEST_P(ExTxtLoadOptions_TrailSpaces, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TrailSpaces(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtLoadOptions_TrailSpaces, ::testing::ValuesIn(ExTxtLoadOptions_TrailSpaces::TestCases()));

} // namespace gtest_test

void ExTxtLoadOptions::DetectDocumentDirection()
{
    //ExStart
    //ExFor:DocumentDirection
    //ExFor:TxtLoadOptions.DocumentDirection
    //ExFor:ParagraphFormat.Bidi
    //ExSummary:Shows how to detect plaintext document text direction.
    // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
    // to modify how we load a plaintext document.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::TxtLoadOptions>();
    
    // Set the "DocumentDirection" property to "DocumentDirection.Auto" automatically detects
    // the direction of every paragraph of text that Aspose.Words loads from plaintext.
    // Each paragraph's "Bidi" property will store its direction.
    loadOptions->set_DocumentDirection(Aspose::Words::Loading::DocumentDirection::Auto);
    
    // Detect Hebrew text as right-to-left.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Hebrew text.txt", loadOptions);
    
    ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Bidi());
    
    // Detect English text as right-to-left.
    doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"English text.txt", loadOptions);
    
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

void ExTxtLoadOptions::AutoNumberingDetection()
{
    //ExStart
    //ExFor:TxtLoadOptions.AutoNumberingDetection
    //ExSummary:Shows how to disable automatic numbering detection.
    auto options = System::MakeObject<Aspose::Words::Loading::TxtLoadOptions>();
    options->set_AutoNumberingDetection(false);
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Number detection.txt", options);
    //ExEnd
    
    int32_t listItemsCount = 0;
    for (auto&& paragraph : System::IterateOver<Aspose::Words::Paragraph>(doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)))
    {
        if (paragraph->get_IsListItem())
        {
            listItemsCount++;
        }
    }
    
    ASSERT_EQ(0, listItemsCount);
}

namespace gtest_test
{

TEST_F(ExTxtLoadOptions, AutoNumberingDetection)
{
    s_instance->AutoNumberingDetection();
}

} // namespace gtest_test

void ExTxtLoadOptions::DetectHyperlinks()
{
    //ExStart:DetectHyperlinks
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:TxtLoadOptions
    //ExFor:TxtLoadOptions.#ctor
    //ExFor:TxtLoadOptions.DetectHyperlinks
    //ExSummary:Shows how to read and display hyperlinks.
    const System::String inputText = System::String(u"Some links in TXT:\n") + u"https://www.aspose.com/\n" + u"https://docs.aspose.com/words/net/\n";
    
    {
        System::SharedPtr<System::IO::Stream> stream = System::MakeObject<System::IO::MemoryStream>();
        System::ArrayPtr<uint8_t> buf = System::Text::Encoding::get_ASCII()->GetBytes(inputText);
        stream->Write(buf, 0, buf->get_Length());
        auto loadOptions = System::MakeObject<Aspose::Words::Loading::TxtLoadOptions>();
        loadOptions->set_DetectHyperlinks(true);
        
        // Load document with hyperlinks.
        auto doc = System::MakeObject<Aspose::Words::Document>(stream, loadOptions);
        
        // Print hyperlinks text.
        for (auto&& field : System::IterateOver(doc->get_Range()->get_Fields()))
        {
            std::cout << field->get_Result() << std::endl;
        }
        
        ASSERT_EQ(doc->get_Range()->get_Fields()->idx_get(0)->get_Result().Trim(), u"https://www.aspose.com/");
        ASSERT_EQ(doc->get_Range()->get_Fields()->idx_get(1)->get_Result().Trim(), u"https://docs.aspose.com/words/net/");
    }
    //ExEnd:DetectHyperlinks
}

namespace gtest_test
{

TEST_F(ExTxtLoadOptions, DetectHyperlinks)
{
    s_instance->DetectHyperlinks();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
