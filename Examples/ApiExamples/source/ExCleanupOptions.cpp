// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExCleanupOptions.h"

using namespace Aspose::Words;
namespace ApiExamples { namespace gtest_test {

class ExCleanupOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExCleanupOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExCleanupOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExCleanupOptions> ExCleanupOptions::s_instance;

TEST_F(ExCleanupOptions, RemoveUnusedResources)
{
    s_instance->RemoveUnusedResources();
}

TEST_F(ExCleanupOptions, RemoveDuplicateStyles)
{
    s_instance->RemoveDuplicateStyles();
}

}} // namespace ApiExamples::gtest_test
