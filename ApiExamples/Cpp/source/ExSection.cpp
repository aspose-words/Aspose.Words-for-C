// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSection.h"

namespace ApiExamples { namespace gtest_test {

class ExSection : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExSection> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExSection>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExSection> ExSection::s_instance;

TEST_F(ExSection, Protect)
{
    s_instance->Protect();
}

TEST_F(ExSection, AddRemove)
{
    s_instance->AddRemove();
}

TEST_F(ExSection, FirstAndLast)
{
    s_instance->FirstAndLast();
}

TEST_F(ExSection, CreateFromScratch)
{
    s_instance->CreateFromScratch();
}

TEST_F(ExSection, EnsureSectionMinimum)
{
    s_instance->EnsureSectionMinimum();
}

TEST_F(ExSection, BodyEnsureMinimum)
{
    s_instance->BodyEnsureMinimum();
}

TEST_F(ExSection, BodyNodeType)
{
    s_instance->BodyNodeType();
}

TEST_F(ExSection, SectionsDeleteAllSections)
{
    s_instance->SectionsDeleteAllSections();
}

TEST_F(ExSection, SectionsAppendSectionContent)
{
    s_instance->SectionsAppendSectionContent();
}

TEST_F(ExSection, SectionsDeleteSectionContent)
{
    s_instance->SectionsDeleteSectionContent();
}

TEST_F(ExSection, SectionsDeleteHeaderFooter)
{
    s_instance->SectionsDeleteHeaderFooter();
}

TEST_F(ExSection, SectionDeleteHeaderFooterShapes)
{
    s_instance->SectionDeleteHeaderFooterShapes();
}

TEST_F(ExSection, SectionsCloneSection)
{
    s_instance->SectionsCloneSection();
}

TEST_F(ExSection, SectionsImportSection)
{
    s_instance->SectionsImportSection();
}

TEST_F(ExSection, MigrateFrom2XImportSection)
{
    s_instance->MigrateFrom2XImportSection();
}

TEST_F(ExSection, ModifyPageSetupInAllSections)
{
    s_instance->ModifyPageSetupInAllSections();
}

TEST_F(ExSection, CultureInfoPageSetupDefaults)
{
    s_instance->CultureInfoPageSetupDefaults();
}

}} // namespace ApiExamples::gtest_test
