// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExCompatibilityOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Settings;
namespace ApiExamples { namespace gtest_test {

class ExCompatibilityOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExCompatibilityOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExCompatibilityOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExCompatibilityOptions> ExCompatibilityOptions::s_instance;

TEST_F(ExCompatibilityOptions, Tables)
{
    s_instance->Tables();
}

TEST_F(ExCompatibilityOptions, Breaks)
{
    s_instance->Breaks();
}

TEST_F(ExCompatibilityOptions, Spacing)
{
    s_instance->Spacing();
}

TEST_F(ExCompatibilityOptions, WordPerfect)
{
    s_instance->WordPerfect();
}

TEST_F(ExCompatibilityOptions, Alignment)
{
    s_instance->Alignment();
}

TEST_F(ExCompatibilityOptions, Legacy)
{
    s_instance->Legacy();
}

TEST_F(ExCompatibilityOptions, List)
{
    s_instance->List();
}

TEST_F(ExCompatibilityOptions, Misc)
{
    s_instance->Misc();
}

}} // namespace ApiExamples::gtest_test
