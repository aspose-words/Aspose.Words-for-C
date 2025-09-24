// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHeaderFooter.h"

#include <testing/test_predicates.h>
#include <system/text/regularexpressions/regex.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/func.h>
#include <system/enumerator_adapter.h>
#include <system/date_time.h>
#include <system/collections/ienumerable.h>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Story.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/ExportHeadersFootersMode.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>


using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(534839304u, ::Aspose::Words::ApiExamples::ExHeaderFooter::ReplaceLog, ThisTypeBaseTypesInfo);

System::String ExHeaderFooter::ReplaceLog::get_Text()
{
    return mTextBuilder->ToString();
}

Aspose::Words::Replacing::ReplaceAction ExHeaderFooter::ReplaceLog::Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args)
{
    mTextBuilder->AppendLine(args->get_MatchNode()->GetText());
    return Aspose::Words::Replacing::ReplaceAction::Skip;
}

ExHeaderFooter::ReplaceLog::ReplaceLog() : mTextBuilder(System::MakeObject<System::Text::StringBuilder>())
{
}


RTTI_INFO_IMPL_HASH(4211736920u, ::Aspose::Words::ApiExamples::ExHeaderFooter, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExHeaderFooter : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExHeaderFooter> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExHeaderFooter>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExHeaderFooter> ExHeaderFooter::s_instance;

} // namespace gtest_test

void ExHeaderFooter::Create()
{
    //ExStart
    //ExFor:HeaderFooter
    //ExFor:HeaderFooter.#ctor(DocumentBase, HeaderFooterType)
    //ExFor:HeaderFooter.HeaderFooterType
    //ExFor:HeaderFooter.IsHeader
    //ExFor:HeaderFooterCollection
    //ExFor:Paragraph.IsEndOfHeaderFooter
    //ExFor:Paragraph.ParentSection
    //ExFor:Paragraph.ParentStory
    //ExFor:Story.AppendParagraph
    //ExSummary:Shows how to create a header and a footer.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create a header and append a paragraph to it. The text in that paragraph
    // will appear at the top of every page of this section, above the main body text.
    auto header = System::MakeObject<Aspose::Words::HeaderFooter>(doc, Aspose::Words::HeaderFooterType::HeaderPrimary);
    doc->get_FirstSection()->get_HeadersFooters()->Add(header);
    
    System::SharedPtr<Aspose::Words::Paragraph> para = header->AppendParagraph(u"My header.");
    
    ASSERT_TRUE(header->get_IsHeader());
    ASSERT_TRUE(para->get_IsEndOfHeaderFooter());
    
    // Create a footer and append a paragraph to it. The text in that paragraph
    // will appear at the bottom of every page of this section, below the main body text.
    auto footer = System::MakeObject<Aspose::Words::HeaderFooter>(doc, Aspose::Words::HeaderFooterType::FooterPrimary);
    doc->get_FirstSection()->get_HeadersFooters()->Add(footer);
    
    para = footer->AppendParagraph(u"My footer.");
    
    ASSERT_FALSE(footer->get_IsHeader());
    ASSERT_TRUE(para->get_IsEndOfHeaderFooter());
    
    ASPOSE_ASSERT_EQ(footer, para->get_ParentStory());
    ASPOSE_ASSERT_EQ(footer->get_ParentSection(), para->get_ParentSection());
    ASPOSE_ASSERT_EQ(footer->get_ParentSection(), header->get_ParentSection());
    
    doc->Save(get_ArtifactsDir() + u"HeaderFooter.Create.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"HeaderFooter.Create.docx");
    
    ASSERT_TRUE(doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->get_Range()->get_Text().Contains(u"My header."));
    ASSERT_TRUE(doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->get_Range()->get_Text().Contains(u"My footer."));
}

namespace gtest_test
{

TEST_F(ExHeaderFooter, Create)
{
    s_instance->Create();
}

} // namespace gtest_test

void ExHeaderFooter::Link()
{
    //ExStart
    //ExFor:HeaderFooter.IsLinkedToPrevious
    //ExFor:HeaderFooterCollection.Item(Int32)
    //ExFor:HeaderFooterCollection.LinkToPrevious(HeaderFooterType,Boolean)
    //ExFor:HeaderFooterCollection.LinkToPrevious(Boolean)
    //ExFor:HeaderFooter.ParentSection
    //ExSummary:Shows how to link headers and footers between sections.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Section 1");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Section 2");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Section 3");
    
    // Move to the first section and create a header and a footer. By default,
    // the header and the footer will only appear on pages in the section that contains them.
    builder->MoveToSection(0);
    
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Write(u"This is the header, which will be displayed in sections 1 and 2.");
    
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterPrimary);
    builder->Write(u"This is the footer, which will be displayed in sections 1, 2 and 3.");
    
    // We can link a section's headers/footers to the previous section's headers/footers
    // to allow the linking section to display the linked section's headers/footers.
    doc->get_Sections()->idx_get(1)->get_HeadersFooters()->LinkToPrevious(true);
    
    // Each section will still have its own header/footer objects. When we link sections,
    // the linking section will display the linked section's header/footers while keeping its own.
    ASPOSE_ASSERT_NE(doc->get_Sections()->idx_get(0)->get_HeadersFooters()->idx_get(0), doc->get_Sections()->idx_get(1)->get_HeadersFooters()->idx_get(0));
    ASPOSE_ASSERT_NE(doc->get_Sections()->idx_get(0)->get_HeadersFooters()->idx_get(0)->get_ParentSection(), doc->get_Sections()->idx_get(1)->get_HeadersFooters()->idx_get(0)->get_ParentSection());
    
    // Link the headers/footers of the third section to the headers/footers of the second section.
    // The second section already links to the first section's header/footers,
    // so linking to the second section will create a link chain.
    // The first, second, and now the third sections will all display the first section's headers.
    doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LinkToPrevious(true);
    
    // We can un-link a previous section's header/footers by passing "false" when calling the LinkToPrevious method.
    doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LinkToPrevious(false);
    
    // We can also select only a specific type of header/footer to link using this method.
    // The third section now will have the same footer as the second and first sections, but not the header.
    doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LinkToPrevious(Aspose::Words::HeaderFooterType::FooterPrimary, true);
    
    // The first section's header/footers cannot link themselves to anything because there is no previous section.
    ASSERT_EQ(2, doc->get_Sections()->idx_get(0)->get_HeadersFooters()->get_Count());
    ASSERT_EQ(2, doc->get_Sections()->idx_get(0)->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> hf)>>([](System::SharedPtr<Aspose::Words::Node> hf) -> bool
    {
        return !(System::ExplicitCast<Aspose::Words::HeaderFooter>(hf))->get_IsLinkedToPrevious();
    }))));
    
    // All the second section's header/footers are linked to the first section's headers/footers.
    ASSERT_EQ(6, doc->get_Sections()->idx_get(1)->get_HeadersFooters()->get_Count());
    ASSERT_EQ(6, doc->get_Sections()->idx_get(1)->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> hf)>>([](System::SharedPtr<Aspose::Words::Node> hf) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::HeaderFooter>(hf))->get_IsLinkedToPrevious();
    }))));
    
    // In the third section, only the footer is linked to the first section's footer via the second section.
    ASSERT_EQ(6, doc->get_Sections()->idx_get(2)->get_HeadersFooters()->get_Count());
    ASSERT_EQ(5, doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> hf)>>([](System::SharedPtr<Aspose::Words::Node> hf) -> bool
    {
        return !(System::ExplicitCast<Aspose::Words::HeaderFooter>(hf))->get_IsLinkedToPrevious();
    }))));
    ASSERT_TRUE(doc->get_Sections()->idx_get(2)->get_HeadersFooters()->idx_get(3)->get_IsLinkedToPrevious());
    
    doc->Save(get_ArtifactsDir() + u"HeaderFooter.Link.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"HeaderFooter.Link.docx");
    
    ASSERT_EQ(2, doc->get_Sections()->idx_get(0)->get_HeadersFooters()->get_Count());
    ASSERT_EQ(2, doc->get_Sections()->idx_get(0)->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> hf)>>([](System::SharedPtr<Aspose::Words::Node> hf) -> bool
    {
        return !(System::ExplicitCast<Aspose::Words::HeaderFooter>(hf))->get_IsLinkedToPrevious();
    }))));
    
    ASSERT_EQ(0, doc->get_Sections()->idx_get(1)->get_HeadersFooters()->get_Count());
    ASSERT_EQ(0, doc->get_Sections()->idx_get(1)->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> hf)>>([](System::SharedPtr<Aspose::Words::Node> hf) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::HeaderFooter>(hf))->get_IsLinkedToPrevious();
    }))));
    
    ASSERT_EQ(5, doc->get_Sections()->idx_get(2)->get_HeadersFooters()->get_Count());
    ASSERT_EQ(5, doc->get_Sections()->idx_get(2)->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> hf)>>([](System::SharedPtr<Aspose::Words::Node> hf) -> bool
    {
        return !(System::ExplicitCast<Aspose::Words::HeaderFooter>(hf))->get_IsLinkedToPrevious();
    }))));
}

namespace gtest_test
{

TEST_F(ExHeaderFooter, Link)
{
    s_instance->Link();
}

} // namespace gtest_test

void ExHeaderFooter::RemoveFooters()
{
    //ExStart
    //ExFor:Section.HeadersFooters
    //ExFor:HeaderFooterCollection
    //ExFor:HeaderFooterCollection.Item(HeaderFooterType)
    //ExFor:HeaderFooter
    //ExSummary:Shows how to delete all footers from a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Header and footer types.docx");
    
    // Iterate through each section and remove footers of every kind.
    for (auto&& section : System::IterateOver(doc->LINQ_OfType<System::SharedPtr<Aspose::Words::Section> >()))
    {
        // There are three kinds of footer and header types.
        // 1 -  The "First" header/footer, which only appears on the first page of a section.
        System::SharedPtr<Aspose::Words::HeaderFooter> footer = section->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterFirst);
        System::SharedPtr<Aspose::Words::HeaderFooter> condExpression = footer;
        if (condExpression != nullptr)
        {
            condExpression->Remove();
        }
        
        // 2 -  The "Primary" header/footer, which appears on odd pages.
        footer = section->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary);
        System::SharedPtr<Aspose::Words::HeaderFooter> condExpression2 = footer;
        if (condExpression2 != nullptr)
        {
            condExpression2->Remove();
        }
        
        // 3 -  The "Even" header/footer, which appears on even pages. 
        footer = section->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterEven);
        System::SharedPtr<Aspose::Words::HeaderFooter> condExpression3 = footer;
        if (condExpression3 != nullptr)
        {
            condExpression3->Remove();
        }
        
        ASSERT_EQ(0, section->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> hf)>>([](System::SharedPtr<Aspose::Words::Node> hf) -> bool
        {
            return !(System::ExplicitCast<Aspose::Words::HeaderFooter>(hf))->get_IsHeader();
        }))));
    }
    
    doc->Save(get_ArtifactsDir() + u"HeaderFooter.RemoveFooters.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"HeaderFooter.RemoveFooters.docx");
    
    ASSERT_EQ(1, doc->get_Sections()->get_Count());
    ASSERT_EQ(0, doc->get_FirstSection()->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> hf)>>([](System::SharedPtr<Aspose::Words::Node> hf) -> bool
    {
        return !(System::ExplicitCast<Aspose::Words::HeaderFooter>(hf))->get_IsHeader();
    }))));
    ASSERT_EQ(3, doc->get_FirstSection()->get_HeadersFooters()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> hf)>>([](System::SharedPtr<Aspose::Words::Node> hf) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::HeaderFooter>(hf))->get_IsHeader();
    }))));
}

namespace gtest_test
{

TEST_F(ExHeaderFooter, RemoveFooters)
{
    s_instance->RemoveFooters();
}

} // namespace gtest_test

void ExHeaderFooter::ExportMode()
{
    //ExStart
    //ExFor:HtmlSaveOptions.ExportHeadersFootersMode
    //ExFor:ExportHeadersFootersMode
    //ExSummary:Shows how to omit headers/footers when saving a document to HTML.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Header and footer types.docx");
    
    // This document contains headers and footers. We can access them via the "HeadersFooters" collection.
    ASSERT_EQ(u"First header", doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderFirst)->GetText().Trim());
    
    // Formats such as .html do not split the document into pages, so headers/footers will not function the same way
    // they would when we open the document as a .docx using Microsoft Word.
    // If we convert a document with headers/footers to html, the conversion will assimilate the headers/footers into body text.
    // We can use a SaveOptions object to omit headers/footers while converting to html.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>(Aspose::Words::SaveFormat::Html);
    saveOptions->set_ExportHeadersFootersMode(Aspose::Words::Saving::ExportHeadersFootersMode::None);
    
    doc->Save(get_ArtifactsDir() + u"HeaderFooter.ExportMode.html", saveOptions);
    
    // Open our saved document and verify that it does not contain the header's text
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"HeaderFooter.ExportMode.html");
    
    ASSERT_FALSE(doc->get_Range()->get_Text().Contains(u"First header"));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExHeaderFooter, ExportMode)
{
    s_instance->ExportMode();
}

} // namespace gtest_test

void ExHeaderFooter::ReplaceText()
{
    //ExStart
    //ExFor:Document.FirstSection
    //ExFor:Section.HeadersFooters
    //ExFor:HeaderFooterCollection.Item(HeaderFooterType)
    //ExFor:HeaderFooter
    //ExFor:Range.Replace(String, String, FindReplaceOptions)
    //ExSummary:Shows how to replace text in a document's footer.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Footer.docx");
    
    System::SharedPtr<Aspose::Words::HeaderFooterCollection> headersFooters = doc->get_FirstSection()->get_HeadersFooters();
    System::SharedPtr<Aspose::Words::HeaderFooter> footer = headersFooters->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary);
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    options->set_MatchCase(false);
    options->set_FindWholeWordsOnly(false);
    
    int32_t currentYear = System::DateTime::get_Now().get_Year();
    footer->get_Range()->Replace(u"(C) 2006 Aspose Pty Ltd.", System::String::Format(u"Copyright (C) {0} by Aspose Pty Ltd.", currentYear), options);
    
    doc->Save(get_ArtifactsDir() + u"HeaderFooter.ReplaceText.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"HeaderFooter.ReplaceText.docx");
    
    ASSERT_TRUE(doc->get_Range()->get_Text().Contains(System::String::Format(u"Copyright (C) {0} by Aspose Pty Ltd.", currentYear)));
}

namespace gtest_test
{

TEST_F(ExHeaderFooter, ReplaceText)
{
    s_instance->ReplaceText();
}

} // namespace gtest_test

void ExHeaderFooter::Order(bool differentFirstPageHeaderFooter)
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Header and footer types.docx");
    
    System::SharedPtr<Aspose::Words::Section> firstPageSection = doc->get_FirstSection();
    
    auto logger = System::MakeObject<Aspose::Words::ApiExamples::ExHeaderFooter::ReplaceLog>();
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>(logger);
    
    // Using a different header/footer for the first page will affect the search order.
    firstPageSection->get_PageSetup()->set_DifferentFirstPageHeaderFooter(differentFirstPageHeaderFooter);
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"(header|footer)"), u"", options);
    
    if (differentFirstPageHeaderFooter)
    {
        ASSERT_EQ(u"First header\nFirst footer\nSecond header\nSecond footer\nThird header\nThird footer\n", logger->get_Text().Replace(u"\r", u""));
    }
    else
    {
        ASSERT_EQ(u"Third header\nFirst header\nThird footer\nFirst footer\nSecond header\nSecond footer\n", logger->get_Text().Replace(u"\r", u""));
    }
}

namespace gtest_test
{

using ExHeaderFooter_Order_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExHeaderFooter::Order)>::type;

struct ExHeaderFooter_Order : public ExHeaderFooter, public Aspose::Words::ApiExamples::ExHeaderFooter, public ::testing::WithParamInterface<ExHeaderFooter_Order_Args>
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

TEST_P(ExHeaderFooter_Order, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Order(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHeaderFooter_Order, ::testing::ValuesIn(ExHeaderFooter_Order::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
