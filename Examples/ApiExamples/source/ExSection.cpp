// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSection.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
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

TEST_F(ExSection, CreateManually)
{
    s_instance->CreateManually();
}

TEST_F(ExSection, EnsureMinimum)
{
    s_instance->EnsureMinimum();
}

TEST_F(ExSection, BodyEnsureMinimum)
{
    s_instance->BodyEnsureMinimum();
}

TEST_F(ExSection, BodyChildNodes)
{
    s_instance->BodyChildNodes();
}

TEST_F(ExSection, Clear)
{
    s_instance->Clear();
}

TEST_F(ExSection, PrependAppendContent)
{
    s_instance->PrependAppendContent();
}

TEST_F(ExSection, ClearContent)
{
    s_instance->ClearContent();
}

TEST_F(ExSection, ClearHeadersFooters)
{
    s_instance->ClearHeadersFooters();
}

TEST_F(ExSection, DeleteHeaderFooterShapes)
{
    s_instance->DeleteHeaderFooterShapes();
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
