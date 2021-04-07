// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSavingCallback.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples { namespace gtest_test {

class ExSavingCallback : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExSavingCallback> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExSavingCallback>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExSavingCallback> ExSavingCallback::s_instance;

TEST_F(ExSavingCallback, CheckThatAllMethodsArePresent)
{
    s_instance->CheckThatAllMethodsArePresent();
}

TEST_F(ExSavingCallback, PageFileNames)
{
    s_instance->PageFileNames();
}

TEST_F(ExSavingCallback, DocumentPartsFileNames)
{
    s_instance->DocumentPartsFileNames();
}

TEST_F(ExSavingCallback, ExternalCssFilenames)
{
    s_instance->ExternalCssFilenames();
}

}} // namespace ApiExamples::gtest_test
