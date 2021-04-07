// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTabStop.h"

using namespace Aspose::Words;
namespace ApiExamples { namespace gtest_test {

class ExTabStop : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExTabStop> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExTabStop>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExTabStop> ExTabStop::s_instance;

TEST_F(ExTabStop, AddTabStops)
{
    s_instance->AddTabStops();
}

TEST_F(ExTabStop, TabStopCollection)
{
    s_instance->TabStopCollection_();
}

TEST_F(ExTabStop, RemoveByIndex)
{
    s_instance->RemoveByIndex();
}

TEST_F(ExTabStop, GetPositionByIndex)
{
    s_instance->GetPositionByIndex();
}

TEST_F(ExTabStop, GetIndexByPosition)
{
    s_instance->GetIndexByPosition();
}

}} // namespace ApiExamples::gtest_test
