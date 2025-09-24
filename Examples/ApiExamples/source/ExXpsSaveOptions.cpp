// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExXpsSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/enumerator_adapter.h>
#include <system/date_time.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Settings/MultiplePagesType.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Saving/XpsSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSet/PageSet.h>
#include <Aspose.Words.Cpp/Model/Saving/OutlineOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/DigitalSignatureDetails.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/CertificateHolder.h>


using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(488843894u, ::Aspose::Words::ApiExamples::ExXpsSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExXpsSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExXpsSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExXpsSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExXpsSaveOptions> ExXpsSaveOptions::s_instance;

} // namespace gtest_test

void ExXpsSaveOptions::OutlineLevels()
{
    //ExStart
    //ExFor:XpsSaveOptions
    //ExFor:XpsSaveOptions.#ctor
    //ExFor:XpsSaveOptions.OutlineOptions
    //ExFor:XpsSaveOptions.SaveFormat
    //ExSummary:Shows how to limit the headings' level that will appear in the outline of a saved XPS document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert headings that can serve as TOC entries of levels 1, 2, and then 3.
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading1);
    
    ASSERT_TRUE(builder->get_ParagraphFormat()->get_IsHeading());
    
    builder->Writeln(u"Heading 1");
    
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading2);
    
    builder->Writeln(u"Heading 1.1");
    builder->Writeln(u"Heading 1.2");
    
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading3);
    
    builder->Writeln(u"Heading 1.2.1");
    builder->Writeln(u"Heading 1.2.2");
    
    // Create an "XpsSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .XPS.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::XpsSaveOptions>();
    
    ASSERT_EQ(Aspose::Words::SaveFormat::Xps, saveOptions->get_SaveFormat());
    
    // The output XPS document will contain an outline, a table of contents that lists headings in the document body.
    // Clicking on an entry in this outline will take us to the location of its respective heading.
    // Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
    // The last two headings we have inserted above will not appear.
    saveOptions->get_OutlineOptions()->set_HeadingsOutlineLevels(2);
    
    doc->Save(get_ArtifactsDir() + u"XpsSaveOptions.OutlineLevels.xps", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExXpsSaveOptions, OutlineLevels)
{
    s_instance->OutlineLevels();
}

} // namespace gtest_test

void ExXpsSaveOptions::BookFold(bool renderTextAsBookFold)
{
    //ExStart
    //ExFor:XpsSaveOptions.#ctor(SaveFormat)
    //ExFor:XpsSaveOptions.UseBookFoldPrintingSettings
    //ExSummary:Shows how to save a document to the XPS format in the form of a book fold.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Paragraphs.docx");
    
    // Create an "XpsSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .XPS.
    auto xpsOptions = System::MakeObject<Aspose::Words::Saving::XpsSaveOptions>(Aspose::Words::SaveFormat::Xps);
    
    // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
    // in the output XPS in a way that helps us use it to make a booklet.
    // Set the "UseBookFoldPrintingSettings" property to "false" to render the XPS normally.
    xpsOptions->set_UseBookFoldPrintingSettings(renderTextAsBookFold);
    
    // If we are rendering the document as a booklet, we must set the "MultiplePages"
    // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
    if (renderTextAsBookFold)
    {
        for (auto&& s : System::IterateOver<Aspose::Words::Section>(doc->get_Sections()))
        {
            s->get_PageSetup()->set_MultiplePages(Aspose::Words::Settings::MultiplePagesType::BookFoldPrinting);
        }
    }
    
    // Once we print this document, we can turn it into a booklet by stacking the pages
    // to come out of the printer and folding down the middle.
    doc->Save(get_ArtifactsDir() + u"XpsSaveOptions.BookFold.xps", xpsOptions);
    //ExEnd
}

namespace gtest_test
{

using ExXpsSaveOptions_BookFold_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExXpsSaveOptions::BookFold)>::type;

struct ExXpsSaveOptions_BookFold : public ExXpsSaveOptions, public Aspose::Words::ApiExamples::ExXpsSaveOptions, public ::testing::WithParamInterface<ExXpsSaveOptions_BookFold_Args>
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

TEST_P(ExXpsSaveOptions_BookFold, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->BookFold(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExXpsSaveOptions_BookFold, ::testing::ValuesIn(ExXpsSaveOptions_BookFold::TestCases()));

} // namespace gtest_test

void ExXpsSaveOptions::ExportExactPages()
{
    //ExStart
    //ExFor:FixedPageSaveOptions.PageSet
    //ExFor:PageSet.#ctor(int[])
    //ExSummary:Shows how to extract pages based on exact page indices.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Add five pages to the document.
    for (int32_t i = 1; i < 6; i++)
    {
        builder->Write(System::String(u"Page ") + i);
        builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    }
    
    // Create an "XpsSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how that method converts the document to .XPS.
    auto xpsOptions = System::MakeObject<Aspose::Words::Saving::XpsSaveOptions>();
    
    // Use the "PageSet" property to select a set of the document's pages to save to output XPS.
    // In this case, we will choose, via a zero-based index, only three pages: page 1, page 2, and page 4.
    xpsOptions->set_PageSet(System::MakeObject<Aspose::Words::Saving::PageSet>(System::MakeArray<int32_t>({0, 1, 3})));
    
    doc->Save(get_ArtifactsDir() + u"XpsSaveOptions.ExportExactPages.xps", xpsOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExXpsSaveOptions, ExportExactPages)
{
    s_instance->ExportExactPages();
}

} // namespace gtest_test

void ExXpsSaveOptions::XpsDigitalSignature()
{
    //ExStart:XpsDigitalSignature
    //GistId:708ce40a68fac5003d46f6b4acfd5ff1
    //ExFor:XpsSaveOptions.DigitalSignatureDetails
    //ExSummary:Shows how to sign XPS document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    System::SharedPtr<Aspose::Words::DigitalSignatures::CertificateHolder> certificateHolder = Aspose::Words::DigitalSignatures::CertificateHolder::Create(get_MyDir() + u"morzal.pfx", u"aw");
    auto options = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    options->set_SignTime(System::DateTime::get_Now());
    options->set_Comments(u"Some comments");
    
    auto digitalSignatureDetails = System::MakeObject<Aspose::Words::Saving::DigitalSignatureDetails>(certificateHolder, options);
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::XpsSaveOptions>();
    saveOptions->set_DigitalSignatureDetails(digitalSignatureDetails);
    
    ASPOSE_ASSERT_EQ(certificateHolder, digitalSignatureDetails->get_CertificateHolder());
    ASSERT_EQ(u"Some comments", digitalSignatureDetails->get_SignOptions()->get_Comments());
    
    doc->Save(get_ArtifactsDir() + u"XpsSaveOptions.XpsDigitalSignature.docx", saveOptions);
    //ExEnd:XpsDigitalSignature
}

namespace gtest_test
{

TEST_F(ExXpsSaveOptions, XpsDigitalSignature)
{
    s_instance->XpsDigitalSignature();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
