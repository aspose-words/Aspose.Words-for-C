// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExNodeImporter.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
namespace ApiExamples { namespace gtest_test {

class ExNodeImporter : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExNodeImporter> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExNodeImporter>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExNodeImporter> ExNodeImporter::s_instance;

using ExNodeImporter_KeepSourceNumbering_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExNodeImporter::KeepSourceNumbering)>::type;

struct ExNodeImporter_KeepSourceNumbering : public ExNodeImporter,
                                            public ApiExamples::ExNodeImporter,
                                            public ::testing::WithParamInterface<ExNodeImporter_KeepSourceNumbering_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExNodeImporter_KeepSourceNumbering, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->KeepSourceNumbering(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExNodeImporter_KeepSourceNumbering, ::testing::ValuesIn(ExNodeImporter_KeepSourceNumbering::TestCases()));

TEST_F(ExNodeImporter, InsertAtBookmark)
{
    s_instance->InsertAtBookmark();
}

TEST_F(ExNodeImporter, InsertAtMergeField)
{
    s_instance->InsertAtMergeField();
}

}} // namespace ApiExamples::gtest_test
