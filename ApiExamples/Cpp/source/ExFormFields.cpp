// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExFormFields.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <system/shared_ptr.h>
#include <system/string.h>
#include <system/text/string_builder.h>

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

TEST_F(ExFormFields, FormFieldsWorkWithProperties)
{
    s_instance->FormFieldsWorkWithProperties();
}

TEST_F(ExFormFields, InsertAndRetrieveFormFields)
{
    s_instance->InsertAndRetrieveFormFields();
}

TEST_F(ExFormFields, DeleteFormField)
{
    s_instance->DeleteFormField();
}

TEST_F(ExFormFields, DeleteFormFieldAssociatedWithBookmark)
{
    s_instance->DeleteFormFieldAssociatedWithBookmark();
}

TEST_F(ExFormFields, FormField)
{
    s_instance->FormField_();
}

}} // namespace ApiExamples::gtest_test
