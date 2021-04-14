// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBorderCollection.h"

using namespace Aspose::Words;
namespace ApiExamples { namespace gtest_test {

class ExBorderCollection : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExBorderCollection> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExBorderCollection>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExBorderCollection> ExBorderCollection::s_instance;

TEST_F(ExBorderCollection, GetBordersEnumerator)
{
    s_instance->GetBordersEnumerator();
}

TEST_F(ExBorderCollection, RemoveAllBorders)
{
    s_instance->RemoveAllBorders();
}

}} // namespace ApiExamples::gtest_test
