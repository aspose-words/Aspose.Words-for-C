// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExViewOptions.h"

#include <testing/test_predicates.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/io/memory_stream.h>
#include <system/io/file.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Settings/ViewType.h>
#include <Aspose.Words.Cpp/Model/Settings/ViewOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>


using namespace Aspose::Words::Settings;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(4137140105u, ::Aspose::Words::ApiExamples::ExViewOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExViewOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExViewOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExViewOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExViewOptions> ExViewOptions::s_instance;

} // namespace gtest_test

void ExViewOptions::SetZoomPercentage()
{
    //ExStart
    //ExFor:Document.ViewOptions
    //ExFor:ViewOptions
    //ExFor:ViewOptions.ViewType
    //ExFor:ViewOptions.ZoomPercent
    //ExFor:ViewOptions.ZoomType
    //ExFor:ZoomType
    //ExFor:ViewType
    //ExSummary:Shows how to set a custom zoom factor, which older versions of Microsoft Word will apply to a document upon loading.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    doc->get_ViewOptions()->set_ViewType(Aspose::Words::Settings::ViewType::PageLayout);
    doc->get_ViewOptions()->set_ZoomPercent(50);
    
    ASSERT_EQ(Aspose::Words::Settings::ZoomType::Custom, doc->get_ViewOptions()->get_ZoomType());
    ASSERT_EQ(Aspose::Words::Settings::ZoomType::None, doc->get_ViewOptions()->get_ZoomType());
    
    doc->Save(get_ArtifactsDir() + u"ViewOptions.SetZoomPercentage.doc");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"ViewOptions.SetZoomPercentage.doc");
    
    ASSERT_EQ(Aspose::Words::Settings::ViewType::PageLayout, doc->get_ViewOptions()->get_ViewType());
    ASPOSE_ASSERT_EQ(50.0, doc->get_ViewOptions()->get_ZoomPercent());
    ASSERT_EQ(Aspose::Words::Settings::ZoomType::None, doc->get_ViewOptions()->get_ZoomType());
}

namespace gtest_test
{

TEST_F(ExViewOptions, SetZoomPercentage)
{
    s_instance->SetZoomPercentage();
}

} // namespace gtest_test

void ExViewOptions::SetZoomType(Aspose::Words::Settings::ZoomType zoomType)
{
    //ExStart
    //ExFor:Document.ViewOptions
    //ExFor:ViewOptions
    //ExFor:ViewOptions.ZoomType
    //ExSummary:Shows how to set a custom zoom type, which older versions of Microsoft Word will apply to a document upon loading.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    // Set the "ZoomType" property to "ZoomType.PageWidth" to get Microsoft Word
    // to automatically zoom the document to fit the width of the page.
    // Set the "ZoomType" property to "ZoomType.FullPage" to get Microsoft Word
    // to automatically zoom the document to make the entire first page visible.
    // Set the "ZoomType" property to "ZoomType.TextFit" to get Microsoft Word
    // to automatically zoom the document to fit the inner text margins of the first page.
    doc->get_ViewOptions()->set_ZoomType(zoomType);
    
    doc->Save(get_ArtifactsDir() + u"ViewOptions.SetZoomType.doc");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"ViewOptions.SetZoomType.doc");
    
    ASSERT_EQ(zoomType, doc->get_ViewOptions()->get_ZoomType());
}

namespace gtest_test
{

using ExViewOptions_SetZoomType_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExViewOptions::SetZoomType)>::type;

struct ExViewOptions_SetZoomType : public ExViewOptions, public Aspose::Words::ApiExamples::ExViewOptions, public ::testing::WithParamInterface<ExViewOptions_SetZoomType_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Settings::ZoomType::PageWidth),
            std::make_tuple(Aspose::Words::Settings::ZoomType::FullPage),
            std::make_tuple(Aspose::Words::Settings::ZoomType::TextFit),
        };
    }
};

TEST_P(ExViewOptions_SetZoomType, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SetZoomType(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExViewOptions_SetZoomType, ::testing::ValuesIn(ExViewOptions_SetZoomType::TestCases()));

} // namespace gtest_test

void ExViewOptions::DisplayBackgroundShape(bool displayBackgroundShape)
{
    //ExStart
    //ExFor:ViewOptions.DisplayBackgroundShape
    //ExSummary:Shows how to hide/display document background images in view options.
    // Use an HTML string to create a new document with a flat background color.
    const System::String html = u"<html>\r\n                <body style='background-color: blue'>\r\n                    <p>Hello world!</p>\r\n                </body>\r\n            </html>";
    
    auto doc = System::MakeObject<Aspose::Words::Document>(System::MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_Unicode()->GetBytes(html)));
    
    // The source for the document has a flat color background,
    // the presence of which will set the "DisplayBackgroundShape" flag to "true".
    ASSERT_TRUE(doc->get_ViewOptions()->get_DisplayBackgroundShape());
    
    // Keep the "DisplayBackgroundShape" as "true" to get the document to display the background color.
    // This may affect some text colors to improve visibility.
    // Set the "DisplayBackgroundShape" to "false" to not display the background color.
    doc->get_ViewOptions()->set_DisplayBackgroundShape(displayBackgroundShape);
    
    doc->Save(get_ArtifactsDir() + u"ViewOptions.DisplayBackgroundShape.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"ViewOptions.DisplayBackgroundShape.docx");
    
    ASPOSE_ASSERT_EQ(displayBackgroundShape, doc->get_ViewOptions()->get_DisplayBackgroundShape());
}

namespace gtest_test
{

using ExViewOptions_DisplayBackgroundShape_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExViewOptions::DisplayBackgroundShape)>::type;

struct ExViewOptions_DisplayBackgroundShape : public ExViewOptions, public Aspose::Words::ApiExamples::ExViewOptions, public ::testing::WithParamInterface<ExViewOptions_DisplayBackgroundShape_Args>
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

TEST_P(ExViewOptions_DisplayBackgroundShape, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DisplayBackgroundShape(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExViewOptions_DisplayBackgroundShape, ::testing::ValuesIn(ExViewOptions_DisplayBackgroundShape::TestCases()));

} // namespace gtest_test

void ExViewOptions::DisplayPageBoundaries(bool doNotDisplayPageBoundaries)
{
    //ExStart
    //ExFor:ViewOptions.DoNotDisplayPageBoundaries
    //ExSummary:Shows how to hide vertical whitespace and headers/footers in view options.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert content that spans across 3 pages.
    builder->Writeln(u"Paragraph 1, Page 1.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Paragraph 2, Page 2.");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Paragraph 3, Page 3.");
    
    // Insert a header and a footer.
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Writeln(u"This is the header.");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterPrimary);
    builder->Writeln(u"This is the footer.");
    
    // This document contains a small amount of content that takes up a few full pages worth of space.
    // Set the "DoNotDisplayPageBoundaries" flag to "true" to get older versions of Microsoft Word to omit headers,
    // footers, and much of the vertical whitespace when displaying our document.
    // Set the "DoNotDisplayPageBoundaries" flag to "false" to get older versions of Microsoft Word
    // to normally display our document.
    doc->get_ViewOptions()->set_DoNotDisplayPageBoundaries(doNotDisplayPageBoundaries);
    
    doc->Save(get_ArtifactsDir() + u"ViewOptions.DisplayPageBoundaries.doc");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"ViewOptions.DisplayPageBoundaries.doc");
    
    ASPOSE_ASSERT_EQ(doNotDisplayPageBoundaries, doc->get_ViewOptions()->get_DoNotDisplayPageBoundaries());
}

namespace gtest_test
{

using ExViewOptions_DisplayPageBoundaries_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExViewOptions::DisplayPageBoundaries)>::type;

struct ExViewOptions_DisplayPageBoundaries : public ExViewOptions, public Aspose::Words::ApiExamples::ExViewOptions, public ::testing::WithParamInterface<ExViewOptions_DisplayPageBoundaries_Args>
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

TEST_P(ExViewOptions_DisplayPageBoundaries, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DisplayPageBoundaries(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExViewOptions_DisplayPageBoundaries, ::testing::ValuesIn(ExViewOptions_DisplayPageBoundaries::TestCases()));

} // namespace gtest_test

void ExViewOptions::FormsDesign(bool useFormsDesign)
{
    //ExStart
    //ExFor:ViewOptions.FormsDesign
    //ExSummary:Shows how to enable/disable forms design mode.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    // Set the "FormsDesign" property to "false" to keep forms design mode disabled.
    // Set the "FormsDesign" property to "true" to enable forms design mode.
    doc->get_ViewOptions()->set_FormsDesign(useFormsDesign);
    
    doc->Save(get_ArtifactsDir() + u"ViewOptions.FormsDesign.xml");
    
    ASPOSE_ASSERT_EQ(useFormsDesign, System::IO::File::ReadAllText(get_ArtifactsDir() + u"ViewOptions.FormsDesign.xml").Contains(u"<w:formsDesign />"));
    //ExEnd
}

namespace gtest_test
{

using ExViewOptions_FormsDesign_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExViewOptions::FormsDesign)>::type;

struct ExViewOptions_FormsDesign : public ExViewOptions, public Aspose::Words::ApiExamples::ExViewOptions, public ::testing::WithParamInterface<ExViewOptions_FormsDesign_Args>
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

TEST_P(ExViewOptions_FormsDesign, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FormsDesign(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExViewOptions_FormsDesign, ::testing::ValuesIn(ExViewOptions_FormsDesign::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
