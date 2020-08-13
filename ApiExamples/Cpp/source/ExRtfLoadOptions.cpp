// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRtfLoadOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Document/RtfLoadOptions.h>
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

class ExRtfLoadOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExRtfLoadOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExRtfLoadOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExRtfLoadOptions> ExRtfLoadOptions::s_instance;

} // namespace gtest_test

void ExRtfLoadOptions::RecognizeUtf8Text(bool doRecognizeUtb8Text)
{
    //ExStart
    //ExFor:RtfLoadOptions
    //ExFor:RtfLoadOptions.#ctor
    //ExFor:RtfLoadOptions.RecognizeUtf8Text
    //ExSummary:Shows how to detect UTF8 characters during import.
    auto loadOptions = MakeObject<RtfLoadOptions>();
    loadOptions->set_RecognizeUtf8Text(doRecognizeUtb8Text);

    auto doc = MakeObject<Document>(MyDir + u"UTF-8 characters.rtf", loadOptions);

    ASSERT_EQ(doRecognizeUtb8Text ? String(u"“John Doe´s list of currency symbols”™\r€, ¢, £, ¥, ¤") : String(u"â€œJohn DoeÂ´s list of currency symbolsâ€\u009dâ„¢\râ‚¬, Â¢, Â£, Â¥, Â¤"), doc->get_FirstSection()->get_Body()->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

using RecognizeUtf8Text_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExRtfLoadOptions::RecognizeUtf8Text)>::type;

struct RecognizeUtf8TextVP : public ExRtfLoadOptions, public ApiExamples::ExRtfLoadOptions, public ::testing::WithParamInterface<RecognizeUtf8Text_Args>
{
    static std::vector<RecognizeUtf8Text_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(RecognizeUtf8TextVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RecognizeUtf8Text(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExRtfLoadOptions, RecognizeUtf8TextVP, ::testing::ValuesIn(RecognizeUtf8TextVP::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
