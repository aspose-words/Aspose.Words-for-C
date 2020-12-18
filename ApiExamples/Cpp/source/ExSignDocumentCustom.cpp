// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSignDocumentCustom.h"

#include <cstdint>
#include <drawing/image.h>
#include <system/array.h>
#include <system/collections/list.h>
#include <system/shared_ptr.h>
#include <system/string.h>

#include "TestData/TestClasses/SignPersonTestClass.h"

using namespace ApiExamples::TestData::TestClasses;
namespace ApiExamples {

System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<TestData::TestClasses::SignPersonTestClass>>> ExSignDocumentCustom::gSignPersonList;

namespace gtest_test {

class ExSignDocumentCustom : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExSignDocumentCustom> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExSignDocumentCustom>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExSignDocumentCustom> ExSignDocumentCustom::s_instance;

TEST_F(ExSignDocumentCustom, Sign)
{
    s_instance->Sign();
}

} // namespace gtest_test

} // namespace ApiExamples
