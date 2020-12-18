// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExFieldOptions.h"

#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <system/globalization/culture_info.h>
#include <system/string.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
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

TEST_F(ExFieldOptions, FieldOptionsCurrentUser)
{
    s_instance->FieldOptionsCurrentUser();
}

TEST_F(ExFieldOptions, FieldOptionsFileName)
{
    s_instance->FieldOptionsFileName();
}

TEST_F(ExFieldOptions, FieldOptionsBidi)
{
    s_instance->FieldOptionsBidi();
}

TEST_F(ExFieldOptions, FieldOptionsLegacyNumberFormat)
{
    s_instance->FieldOptionsLegacyNumberFormat();
}

TEST_F(ExFieldOptions, FieldOptionsPreProcessCulture)
{
    s_instance->FieldOptionsPreProcessCulture();
}

TEST_F(ExFieldOptions, FieldOptionsToaCategories)
{
    s_instance->FieldOptionsToaCategories();
}

TEST_F(ExFieldOptions, FieldOptionsUseInvariantCultureNumberFormat)
{
    s_instance->FieldOptionsUseInvariantCultureNumberFormat();
}

TEST_F(ExFieldOptions, DefineDateTimeFormatting)
{
    s_instance->DefineDateTimeFormatting();
}

}} // namespace ApiExamples::gtest_test
