// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExVbaProject.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Vba;
namespace ApiExamples { namespace gtest_test {

class ExVbaProject : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExVbaProject> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExVbaProject>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExVbaProject> ExVbaProject::s_instance;

TEST_F(ExVbaProject, CreateNewVbaProject)
{
    s_instance->CreateNewVbaProject();
}

TEST_F(ExVbaProject, CloneVbaProject)
{
    s_instance->CloneVbaProject();
}

TEST_F(ExVbaProject, RemoveVbaReference)
{
    s_instance->RemoveVbaReference();
}

}} // namespace ApiExamples::gtest_test
