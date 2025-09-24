// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExPsSaveOptions.h"

#include <system/test_tools/method_argument_tuple.h>
#include <system/string.h>
#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Settings/MultiplePagesType.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/PsSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3882588110u, ::Aspose::Words::ApiExamples::ExPsSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExPsSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExPsSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExPsSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExPsSaveOptions> ExPsSaveOptions::s_instance;

} // namespace gtest_test

void ExPsSaveOptions::UseBookFoldPrintingSettings(bool renderTextAsBookFold)
{
    //ExStart
    //ExFor:PsSaveOptions
    //ExFor:PsSaveOptions.SaveFormat
    //ExFor:PsSaveOptions.UseBookFoldPrintingSettings
    //ExSummary:Shows how to save a document to the Postscript format in the form of a book fold.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Paragraphs.docx");
    
    // Create a "PsSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to PostScript.
    // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
    // in the output Postscript document in a way that helps us make a booklet out of it.
    // Set the "UseBookFoldPrintingSettings" property to "false" to save the document normally.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::PsSaveOptions>();
    saveOptions->set_SaveFormat(Aspose::Words::SaveFormat::Ps);
    saveOptions->set_UseBookFoldPrintingSettings(renderTextAsBookFold);
    
    // If we are rendering the document as a booklet, we must set the "MultiplePages"
    // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
    for (auto&& s : System::IterateOver<Aspose::Words::Section>(doc->get_Sections()))
    {
        s->get_PageSetup()->set_MultiplePages(Aspose::Words::Settings::MultiplePagesType::BookFoldPrinting);
    }
    
    // Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
    // and the contents will line up in a way that creates a booklet.
    doc->Save(get_ArtifactsDir() + u"PsSaveOptions.UseBookFoldPrintingSettings.ps", saveOptions);
    //ExEnd
}

namespace gtest_test
{

using ExPsSaveOptions_UseBookFoldPrintingSettings_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExPsSaveOptions::UseBookFoldPrintingSettings)>::type;

struct ExPsSaveOptions_UseBookFoldPrintingSettings : public ExPsSaveOptions, public Aspose::Words::ApiExamples::ExPsSaveOptions, public ::testing::WithParamInterface<ExPsSaveOptions_UseBookFoldPrintingSettings_Args>
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

TEST_P(ExPsSaveOptions_UseBookFoldPrintingSettings, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UseBookFoldPrintingSettings(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPsSaveOptions_UseBookFoldPrintingSettings, ::testing::ValuesIn(ExPsSaveOptions_UseBookFoldPrintingSettings::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
