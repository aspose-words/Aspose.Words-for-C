// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExFormFields.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
namespace ApiExamples { namespace gtest_test {

class ExFormFields : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExFormFields> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExFormFields>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExFormFields> ExFormFields::s_instance;

TEST_F(ExFormFields, Create)
{
    s_instance->Create();
}

TEST_F(ExFormFields, TextInput)
{
    s_instance->TextInput();
}

TEST_F(ExFormFields, DeleteFormField)
{
    s_instance->DeleteFormField();
}

TEST_F(ExFormFields, DeleteFormFieldAssociatedWithBookmark)
{
    s_instance->DeleteFormFieldAssociatedWithBookmark();
}

TEST_F(ExFormFields, FormFieldFontFormatting)
{
    s_instance->FormFieldFontFormatting();
}

TEST_F(ExFormFields, Visitor)
{
    s_instance->Visitor();
}

TEST_F(ExFormFields, DropDownItemCollection)
{
    s_instance->DropDownItemCollection_();
}

}} // namespace ApiExamples::gtest_test
