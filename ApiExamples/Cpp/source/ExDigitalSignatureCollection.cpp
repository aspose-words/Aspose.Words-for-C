// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDigitalSignatureCollection.h"

namespace ApiExamples { namespace gtest_test {

class ExDigitalSignatureCollection : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExDigitalSignatureCollection> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExDigitalSignatureCollection>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExDigitalSignatureCollection> ExDigitalSignatureCollection::s_instance;

TEST_F(ExDigitalSignatureCollection, GetEnumerator)
{
    s_instance->GetEnumerator();
}

}} // namespace ApiExamples::gtest_test
