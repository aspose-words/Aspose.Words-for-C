// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExFieldOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Loading;
namespace ApiExamples { namespace gtest_test {

class ExFieldOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExFieldOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExFieldOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExFieldOptions> ExFieldOptions::s_instance;

TEST_F(ExFieldOptions, CurrentUser)
{
    s_instance->CurrentUser();
}

TEST_F(ExFieldOptions, FileName)
{
    s_instance->FileName();
}

TEST_F(ExFieldOptions, Bidi)
{
    s_instance->Bidi();
}

TEST_F(ExFieldOptions, LegacyNumberFormat)
{
    s_instance->LegacyNumberFormat();
}

TEST_F(ExFieldOptions, PreProcessCulture)
{
    s_instance->PreProcessCulture();
}

TEST_F(ExFieldOptions, TableOfAuthorityCategories)
{
    s_instance->TableOfAuthorityCategories();
}

TEST_F(ExFieldOptions, UseInvariantCultureNumberFormat)
{
    s_instance->UseInvariantCultureNumberFormat();
}

TEST_F(ExFieldOptions, DefineDateTimeFormatting)
{
    s_instance->DefineDateTimeFormatting();
}

}} // namespace ApiExamples::gtest_test
