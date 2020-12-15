// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExPsSaveOptions.h"

namespace ApiExamples { namespace gtest_test {

class ExPsSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExPsSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExPsSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExPsSaveOptions> ExPsSaveOptions::s_instance;

TEST_F(ExPsSaveOptions, UseBookFoldPrintingSettings)
{
    s_instance->UseBookFoldPrintingSettings();
}

}} // namespace ApiExamples::gtest_test
