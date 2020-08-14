// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExXpsSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/io/file_info.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/XpsSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples {

namespace gtest_test
{

class ExXpsSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExXpsSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExXpsSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExXpsSaveOptions> ExXpsSaveOptions::s_instance;

} // namespace gtest_test

void ExXpsSaveOptions::OptimizeOutput(bool optimizeOutput)
{
    //ExStart
    //ExFor:FixedPageSaveOptions.OptimizeOutput
    //ExSummary:Shows how to optimize document objects while saving to xps.
    auto doc = MakeObject<Document>(MyDir + u"Unoptimized document.docx");

    // When saving to .xps, we can use SaveOptions to optimize the output in some cases
    auto saveOptions = MakeObject<XpsSaveOptions>();
    saveOptions->set_OptimizeOutput(optimizeOutput);

    doc->Save(ArtifactsDir + u"XpsSaveOptions.OptimizeOutput.xps", saveOptions);

    // The input document had adjacent runs with the same formatting, which, if output optimization was enabled,
    // have been combined to save space
    auto outFileInfo = MakeObject<System::IO::FileInfo>(ArtifactsDir + u"XpsSaveOptions.OptimizeOutput.xps");

    if (optimizeOutput)
    {
        ASSERT_TRUE(outFileInfo->get_Length() < 45000);
    }
    else
    {
        ASSERT_TRUE(outFileInfo->get_Length() > 60000);
    }
    //ExEnd
}

namespace gtest_test
{

using OptimizeOutput_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExXpsSaveOptions::OptimizeOutput)>::type;

struct OptimizeOutputVP : public ExXpsSaveOptions, public ApiExamples::ExXpsSaveOptions, public ::testing::WithParamInterface<OptimizeOutput_Args>
{
    static std::vector<OptimizeOutput_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(OptimizeOutputVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->OptimizeOutput(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExXpsSaveOptions, OptimizeOutputVP, ::testing::ValuesIn(OptimizeOutputVP::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
