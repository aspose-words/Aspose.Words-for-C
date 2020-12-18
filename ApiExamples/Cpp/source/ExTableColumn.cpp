// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTableColumn.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <system/collections/list.h>
#include <system/shared_ptr.h>
#include <system/string.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;
namespace ApiExamples { namespace gtest_test {

class ExTableColumn : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExTableColumn> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExTableColumn>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExTableColumn> ExTableColumn::s_instance;

TEST_F(ExTableColumn, RemoveColumnFromTable)
{
    s_instance->RemoveColumnFromTable();
}

TEST_F(ExTableColumn, Insert)
{
    s_instance->Insert();
}

TEST_F(ExTableColumn, TableColumnToTxt)
{
    s_instance->TableColumnToTxt();
}

}} // namespace ApiExamples::gtest_test
