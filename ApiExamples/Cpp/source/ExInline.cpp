// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExInline.h"

namespace ApiExamples { namespace gtest_test {

class ExInline : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExInline> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExInline>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExInline> ExInline::s_instance;

TEST_F(ExInline, InlineRevisions)
{
    s_instance->InlineRevisions();
}

}} // namespace ApiExamples::gtest_test
