// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExLists.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Lists;
namespace ApiExamples { namespace gtest_test {

class ExLists : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExLists> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExLists>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExLists> ExLists::s_instance;

TEST_F(ExLists, ApplyDefaultBulletsAndNumbers)
{
    s_instance->ApplyDefaultBulletsAndNumbers();
}

TEST_F(ExLists, SpecifyListLevel)
{
    s_instance->SpecifyListLevel();
}

TEST_F(ExLists, NestedLists)
{
    s_instance->NestedLists();
}

TEST_F(ExLists, CreateCustomList)
{
    s_instance->CreateCustomList();
}

TEST_F(ExLists, RestartNumberingUsingListCopy)
{
    s_instance->RestartNumberingUsingListCopy();
}

TEST_F(ExLists, CreateAndUseListStyle)
{
    s_instance->CreateAndUseListStyle();
}

TEST_F(ExLists, DetectBulletedParagraphs)
{
    s_instance->DetectBulletedParagraphs();
}

TEST_F(ExLists, RemoveBulletsFromParagraphs)
{
    s_instance->RemoveBulletsFromParagraphs();
}

TEST_F(ExLists, ApplyExistingListToParagraphs)
{
    s_instance->ApplyExistingListToParagraphs();
}

TEST_F(ExLists, ApplyNewListToParagraphs)
{
    s_instance->ApplyNewListToParagraphs();
}

TEST_F(ExLists, OutlineHeadingTemplates)
{
    s_instance->OutlineHeadingTemplates();
}

TEST_F(ExLists, PrintOutAllLists)
{
    s_instance->PrintOutAllLists();
}

TEST_F(ExLists, ListDocument)
{
    s_instance->ListDocument();
}

TEST_F(ExLists, CreateListRestartAfterHigher)
{
    s_instance->CreateListRestartAfterHigher();
}

TEST_F(ExLists, GetListLabels)
{
    s_instance->GetListLabels();
}

TEST_F(ExLists, IgnoreOnJenkins_CreatePictureBullet)
{
    RecordProperty("category", "IgnoreOnJenkins");

    s_instance->CreatePictureBullet();
}

}} // namespace ApiExamples::gtest_test
