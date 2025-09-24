// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRtfLoadOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Loading/RtfLoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Loading;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3862892644u, ::Aspose::Words::ApiExamples::ExRtfLoadOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExRtfLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExRtfLoadOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExRtfLoadOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExRtfLoadOptions> ExRtfLoadOptions::s_instance;

} // namespace gtest_test

void ExRtfLoadOptions::RecognizeUtf8Text(bool recognizeUtf8Text)
{
    //ExStart
    //ExFor:RtfLoadOptions
    //ExFor:RtfLoadOptions.#ctor
    //ExFor:RtfLoadOptions.RecognizeUtf8Text
    //ExSummary:Shows how to detect UTF-8 characters while loading an RTF document.
    // Create an "RtfLoadOptions" object to modify how we load an RTF document.
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::RtfLoadOptions>();
    
    // Set the "RecognizeUtf8Text" property to "false" to assume that the document uses the ISO 8859-1 charset
    // and loads every character in the document.
    // Set the "RecognizeUtf8Text" property to "true" to parse any variable-length characters that may occur in the text.
    loadOptions->set_RecognizeUtf8Text(recognizeUtf8Text);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"UTF-8 characters.rtf", loadOptions);
    
    ASSERT_EQ(recognizeUtf8Text ? System::String(u"“John Doe´s list of currency symbols”™\r") + u"€, ¢, £, ¥, ¤" : System::String(u"â€œJohn DoeÂ´s list of currency symbolsâ€\u009dâ„¢\r") + u"â‚¬, Â¢, Â£, Â¥, Â¤", doc->get_FirstSection()->get_Body()->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

using ExRtfLoadOptions_RecognizeUtf8Text_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExRtfLoadOptions::RecognizeUtf8Text)>::type;

struct ExRtfLoadOptions_RecognizeUtf8Text : public ExRtfLoadOptions, public Aspose::Words::ApiExamples::ExRtfLoadOptions, public ::testing::WithParamInterface<ExRtfLoadOptions_RecognizeUtf8Text_Args>
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

TEST_P(ExRtfLoadOptions_RecognizeUtf8Text, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RecognizeUtf8Text(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExRtfLoadOptions_RecognizeUtf8Text, ::testing::ValuesIn(ExRtfLoadOptions_RecognizeUtf8Text::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
