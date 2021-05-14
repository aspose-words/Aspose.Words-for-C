// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHyphenation.h"

using namespace Aspose::Words;
namespace ApiExamples { namespace gtest_test {

class ExHyphenation : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExHyphenation> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExHyphenation>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExHyphenation> ExHyphenation::s_instance;

TEST_F(ExHyphenation, Dictionary)
{
    s_instance->Dictionary();
}

TEST_F(ExHyphenation, RegisterDictionary)
{
    s_instance->RegisterDictionary();
}

}} // namespace ApiExamples::gtest_test
