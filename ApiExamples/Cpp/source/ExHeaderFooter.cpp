// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHeaderFooter.h"

#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <system/shared_ptr.h>
#include <system/text/string_builder.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;
namespace ApiExamples { namespace gtest_test {

class ExHeaderFooter : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExHeaderFooter> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExHeaderFooter>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExHeaderFooter> ExHeaderFooter::s_instance;

TEST_F(ExHeaderFooter, HeaderFooterCreate)
{
    s_instance->HeaderFooterCreate();
}

TEST_F(ExHeaderFooter, HeaderFooterLink)
{
    s_instance->HeaderFooterLink();
}

TEST_F(ExHeaderFooter, RemoveFooters)
{
    s_instance->RemoveFooters();
}

TEST_F(ExHeaderFooter, SetExportHeadersFootersMode)
{
    s_instance->SetExportHeadersFootersMode();
}

TEST_F(ExHeaderFooter, ReplaceText)
{
    s_instance->ReplaceText();
}

TEST_F(ExHeaderFooter, HeaderFooterOrder)
{
    s_instance->HeaderFooterOrder();
}

TEST_F(ExHeaderFooter, Primer)
{
    s_instance->Primer();
}

}} // namespace ApiExamples::gtest_test
