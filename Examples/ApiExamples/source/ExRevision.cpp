// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRevision.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
namespace ApiExamples { namespace gtest_test {

class ExRevision : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExRevision> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExRevision>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExRevision> ExRevision::s_instance;

TEST_F(ExRevision, Revisions)
{
    s_instance->Revisions();
}

TEST_F(ExRevision, RevisionCollection)
{
    s_instance->RevisionCollection_();
}

TEST_F(ExRevision, GetInfoAboutRevisionsInRevisionGroups)
{
    s_instance->GetInfoAboutRevisionsInRevisionGroups();
}

TEST_F(ExRevision, GetSpecificRevisionGroup)
{
    s_instance->GetSpecificRevisionGroup();
}

TEST_F(ExRevision, ShowRevisionBalloons)
{
    s_instance->ShowRevisionBalloons();
}

TEST_F(ExRevision, RevisionOptions)
{
    s_instance->RevisionOptions_();
}

}} // namespace ApiExamples::gtest_test
