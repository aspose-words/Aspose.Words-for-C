// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExViewOptions.h"

#include <testing/test_predicates.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/io/memory_stream.h>
#include <system/io/file.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Settings/ZoomType.h>
#include <Aspose.Words.Cpp/Model/Settings/ViewType.h>
#include <Aspose.Words.Cpp/Model/Settings/ViewOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Saving/WordML2003SaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace ApiExamples {

namespace gtest_test
{

class ExViewOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExViewOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExViewOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExViewOptions> ExViewOptions::s_instance;

} // namespace gtest_test

void ExViewOptions::SetZoom()
{
    //ExStart
    //ExFor:Document.ViewOptions
    //ExFor:ViewOptions
    //ExFor:ViewOptions.ViewType
    //ExFor:ViewOptions.ZoomType
    //ExFor:ViewOptions.ZoomPercent
    //ExFor:ViewType
    //ExSummary:Shows how to make sure the document is displayed at 50% zoom when opened in Microsoft Word.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    // We can set the zoom factor to a percentage
    doc->get_ViewOptions()->set_ViewType(Aspose::Words::Settings::ViewType::PageLayout);
    doc->get_ViewOptions()->set_ZoomPercent(50);

    // Or we can set the ZoomType to a different value to avoid using percentages
    ASSERT_EQ(Aspose::Words::Settings::ZoomType::None, doc->get_ViewOptions()->get_ZoomType());

    doc->Save(ArtifactsDir + u"ViewOptions.SetZoom.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"ViewOptions.SetZoom.docx");

    ASSERT_EQ(Aspose::Words::Settings::ViewType::PageLayout, doc->get_ViewOptions()->get_ViewType());
    ASPOSE_ASSERT_EQ(50.0, doc->get_ViewOptions()->get_ZoomPercent());
    ASSERT_EQ(Aspose::Words::Settings::ZoomType::None, doc->get_ViewOptions()->get_ZoomType());
}

namespace gtest_test
{

TEST_F(ExViewOptions, SetZoom)
{
    s_instance->SetZoom();
}

} // namespace gtest_test

void ExViewOptions::DisplayBackgroundShape()
{
    //ExStart
    //ExFor:ViewOptions.DisplayBackgroundShape
    //ExSummary:Shows how to hide/display document background images in view options.
    // Create a new document from an html string with a flat background color
    const String html = u"<html>\r\n                <body style='background-color: blue'>\r\n                    <p>Hello world!</p>\r\n                </body>\r\n            </html>";

    auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_Unicode()->GetBytes(html)));

    // The source for the document has a flat color background, the presence of which will turn on the DisplayBackgroundShape flag
    // We can disable it like this
    doc->get_ViewOptions()->set_DisplayBackgroundShape(false);

    doc->Save(ArtifactsDir + u"ViewOptions.DisplayBackgroundShape.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"ViewOptions.DisplayBackgroundShape.docx");

    ASSERT_FALSE(doc->get_ViewOptions()->get_DisplayBackgroundShape());
}

namespace gtest_test
{

TEST_F(ExViewOptions, DisplayBackgroundShape)
{
    s_instance->DisplayBackgroundShape();
}

} // namespace gtest_test

void ExViewOptions::DisplayPageBoundaries()
{
    //ExStart
    //ExFor:ViewOptions.DoNotDisplayPageBoundaries
    //ExSummary:Shows how to hide vertical whitespace and headers/footers in view options.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert content spanning 3 pages
    builder->Writeln(u"Paragraph 1, Page 1");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Paragraph 2, Page 2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Paragraph 3, Page 3");

    // Insert a header and a footer
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Writeln(u"Header");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterPrimary);
    builder->Writeln(u"Footer");

    // In this case we have a lot of space taken up by quite a little amount of content
    // In older versions of Microsoft Word, we can hide headers/footers and compact vertical whitespace of pages
    // to give the document's main body content some flow by setting this flag
    doc->get_ViewOptions()->set_DoNotDisplayPageBoundaries(true);

    doc->Save(ArtifactsDir + u"ViewOptions.DisplayPageBoundaries.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"ViewOptions.DisplayPageBoundaries.docx");

    ASSERT_TRUE(doc->get_ViewOptions()->get_DoNotDisplayPageBoundaries());
}

namespace gtest_test
{

TEST_F(ExViewOptions, DisplayPageBoundaries)
{
    s_instance->DisplayPageBoundaries();
}

} // namespace gtest_test

void ExViewOptions::FormsDesign(bool useFormsDesign)
{
    //ExStart
    //ExFor:ViewOptions.FormsDesign
    //ExFor:WordML2003SaveOptions
    //ExFor:WordML2003SaveOptions.SaveFormat
    //ExSummary:Shows how to save to a .wml document while applying save options.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    auto options = MakeObject<WordML2003SaveOptions>();
    options->set_SaveFormat(Aspose::Words::SaveFormat::WordML);
    options->set_MemoryOptimization(true);
    options->set_PrettyFormat(true);

    // Enables forms design mode in WordML documents
    doc->get_ViewOptions()->set_FormsDesign(useFormsDesign);

    doc->Save(ArtifactsDir + u"ViewOptions.FormsDesign.xml", options);

    ASPOSE_ASSERT_EQ(useFormsDesign, System::IO::File::ReadAllText(ArtifactsDir + u"ViewOptions.FormsDesign.xml").Contains(u"<w:formsDesign />"));
    //ExEnd
}

namespace gtest_test
{

using FormsDesign_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExViewOptions::FormsDesign)>::type;

struct FormsDesignVP : public ExViewOptions, public ApiExamples::ExViewOptions, public ::testing::WithParamInterface<FormsDesign_Args>
{
    static std::vector<FormsDesign_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(FormsDesignVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FormsDesign(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExViewOptions, FormsDesignVP, ::testing::ValuesIn(FormsDesignVP::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
