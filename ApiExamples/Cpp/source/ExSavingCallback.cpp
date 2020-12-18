// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSavingCallback.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/CssSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentPartSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentSplitCriteria.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSavingArgs.h>
#include <system/enum_helpers.h>
#include <system/string.h>

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

TEST_F(ExSavingCallback, PageFileName)
{
    s_instance->PageFileName();
}

TEST_F(ExSavingCallback, DocumentParts)
{
    s_instance->DocumentParts();
}

TEST_F(ExSavingCallback, CssSavingCallback)
{
    s_instance->CssSavingCallback();
}

}} // namespace ApiExamples::gtest_test
