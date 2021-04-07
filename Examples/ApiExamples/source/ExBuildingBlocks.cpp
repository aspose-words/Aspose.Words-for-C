// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBuildingBlocks.h"

using namespace Aspose::Words;
using namespace Aspose::Words::BuildingBlocks;
namespace ApiExamples { namespace gtest_test {

class ExBuildingBlocks : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExBuildingBlocks> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExBuildingBlocks>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExBuildingBlocks> ExBuildingBlocks::s_instance;

TEST_F(ExBuildingBlocks, CreateAndInsert)
{
    s_instance->CreateAndInsert();
}

TEST_F(ExBuildingBlocks, GlossaryDocument)
{
    s_instance->GlossaryDocument_();
}

}} // namespace ApiExamples::gtest_test
