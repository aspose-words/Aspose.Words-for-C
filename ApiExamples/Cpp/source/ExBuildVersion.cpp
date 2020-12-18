// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBuildVersion.h"

namespace ApiExamples { namespace gtest_test {

class ExBuildVersion : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExBuildVersion> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExBuildVersion>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExBuildVersion> ExBuildVersion::s_instance;

TEST_F(ExBuildVersion, PrintBuildVersionInfo)
{
    s_instance->PrintBuildVersionInfo();
}

}} // namespace ApiExamples::gtest_test
