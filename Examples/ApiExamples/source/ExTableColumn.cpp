// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTableColumn.h"

#include <testing/test_predicates.h>
#include <system/text/string_builder.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/collections/ienumerable.h>
#include <iostream>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(427892106u, ::Aspose::Words::ApiExamples::ExTableColumn::Column, ThisTypeBaseTypesInfo);

System::ArrayPtr<System::SharedPtr<Aspose::Words::Tables::Cell>> ExTableColumn::Column::get_Cells()
{
    return GetColumnCells()->ToArray();
}

ExTableColumn::Column::Column(System::SharedPtr<Aspose::Words::Tables::Table> table, int32_t columnIndex)
    : mColumnIndex(0)
{
    mTable = table;
    if (mTable == nullptr)
    {
        throw System::ArgumentException(u"table");
    }
    
    mColumnIndex = columnIndex;
}

MEMBER_FUNCTION_MAKE_OBJECT_DEFINITION(ExTableColumn::Column, CODEPORTING_ARGS(System::SharedPtr<Aspose::Words::Tables::Table> table, int32_t columnIndex), CODEPORTING_ARGS(table,columnIndex));

System::SharedPtr<Aspose::Words::ApiExamples::ExTableColumn::Column> ExTableColumn::Column::FromIndex(System::SharedPtr<Aspose::Words::Tables::Table> table, int32_t columnIndex)
{
    return Aspose::Words::ApiExamples::ExTableColumn::Column::MakeObject(table, columnIndex);
}

int32_t ExTableColumn::Column::IndexOf(System::SharedPtr<Aspose::Words::Tables::Cell> cell)
{
    return GetColumnCells()->IndexOf(cell);
}

System::SharedPtr<Aspose::Words::ApiExamples::ExTableColumn::Column> ExTableColumn::Column::InsertColumnBefore()
{
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Tables::Cell>> columnCells = get_Cells();
    
    if (columnCells->get_Length() == 0)
    {
        throw System::ArgumentException(u"Column must not be empty");
    }
    
    // Create a clone of this column
    for (System::SharedPtr<Aspose::Words::Tables::Cell> cell : columnCells)
    {
        cell->get_ParentRow()->InsertBefore<System::SharedPtr<Aspose::Words::Node>>(System::ExplicitCast<Aspose::Words::Node>(cell)->Clone(false), cell);
    }
    
    
    auto newColumn = Aspose::Words::ApiExamples::ExTableColumn::Column::MakeObject(columnCells[0]->get_ParentRow()->get_ParentTable(), mColumnIndex);
    
    // We want to make sure that the cells are all valid to work with (have at least one paragraph).
    for (System::SharedPtr<Aspose::Words::Tables::Cell> cell : newColumn->get_Cells())
    {
        cell->EnsureMinimum();
    }
    
    
    // Increment the index of this column represents since there is a new column before it.
    mColumnIndex++;
    
    return newColumn;
}

void ExTableColumn::Column::Remove()
{
    for (System::SharedPtr<Aspose::Words::Tables::Cell> cell : get_Cells())
    {
        cell->Remove();
    }
    
}

System::String ExTableColumn::Column::ToTxt()
{
    System::SharedPtr<System::Text::StringBuilder> builder = System::MakeObject<System::Text::StringBuilder>();
    
    for (System::SharedPtr<Aspose::Words::Tables::Cell> cell : get_Cells())
    {
        builder->Append(cell->ToString(Aspose::Words::SaveFormat::Text));
    }
    
    
    return builder->ToString();
}

System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Tables::Cell>>> ExTableColumn::Column::GetColumnCells()
{
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Tables::Cell>>> columnCells = System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Tables::Cell>>>();
    
    for (auto&& row : System::IterateOver(mTable->get_Rows()->LINQ_OfType<System::SharedPtr<Aspose::Words::Tables::Row> >()))
    {
        System::SharedPtr<Aspose::Words::Tables::Cell> cell = row->get_Cells()->idx_get(mColumnIndex);
        if (cell != nullptr)
        {
            columnCells->Add(cell);
        }
    }
    
    return columnCells;
}


RTTI_INFO_IMPL_HASH(3765276756u, ::Aspose::Words::ApiExamples::ExTableColumn, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExTableColumn : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExTableColumn> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExTableColumn>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExTableColumn> ExTableColumn::s_instance;

} // namespace gtest_test

void ExTableColumn::RemoveColumnFromTable()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    auto table = System::ExplicitCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));
    
    System::SharedPtr<Aspose::Words::ApiExamples::ExTableColumn::Column> column = Aspose::Words::ApiExamples::ExTableColumn::Column::FromIndex(table, 2);
    column->Remove();
    
    doc->Save(get_ArtifactsDir() + u"TableColumn.RemoveColumn.doc");
    
    ASSERT_EQ(16, table->GetChildNodes(Aspose::Words::NodeType::Cell, true)->get_Count());
    ASSERT_EQ(u"Cell 7 contents", table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2)->ToString(Aspose::Words::SaveFormat::Text).Trim());
    ASSERT_EQ(u"Cell 11 contents", table->get_LastRow()->get_Cells()->idx_get(2)->ToString(Aspose::Words::SaveFormat::Text).Trim());
}

namespace gtest_test
{

TEST_F(ExTableColumn, RemoveColumnFromTable)
{
    s_instance->RemoveColumnFromTable();
}

} // namespace gtest_test

void ExTableColumn::Insert()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    auto table = System::ExplicitCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));
    
    System::SharedPtr<Aspose::Words::ApiExamples::ExTableColumn::Column> column = Aspose::Words::ApiExamples::ExTableColumn::Column::FromIndex(table, 1);
    
    // Create a new column to the left of this column.
    // This is the same as using the "Insert Column Before" command in Microsoft Word.
    System::SharedPtr<Aspose::Words::ApiExamples::ExTableColumn::Column> newColumn = column->InsertColumnBefore();
    
    // Add some text to each cell in the column.
    for (System::SharedPtr<Aspose::Words::Tables::Cell> cell : newColumn->get_Cells())
    {
        cell->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, System::String(u"Column Text ") + newColumn->IndexOf(cell)));
    }
    
    
    doc->Save(get_ArtifactsDir() + u"TableColumn.Insert.doc");
    
    ASSERT_EQ(24, table->GetChildNodes(Aspose::Words::NodeType::Cell, true)->get_Count());
    ASSERT_EQ(u"Column Text 0", table->get_FirstRow()->get_Cells()->idx_get(1)->ToString(Aspose::Words::SaveFormat::Text).Trim());
    ASSERT_EQ(u"Column Text 3", table->get_LastRow()->get_Cells()->idx_get(1)->ToString(Aspose::Words::SaveFormat::Text).Trim());
}

namespace gtest_test
{

TEST_F(ExTableColumn, Insert)
{
    s_instance->Insert();
}

} // namespace gtest_test

void ExTableColumn::TableColumnToTxt()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Tables.docx");
    auto table = System::ExplicitCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));
    
    System::SharedPtr<Aspose::Words::ApiExamples::ExTableColumn::Column> column = Aspose::Words::ApiExamples::ExTableColumn::Column::FromIndex(table, 0);
    std::cout << column->ToTxt() << std::endl;
    
    ASSERT_EQ(u"\rRow 1\rRow 2\rRow 3\r", column->ToTxt());
}

namespace gtest_test
{

TEST_F(ExTableColumn, TableColumnToTxt)
{
    s_instance->TableColumnToTxt();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
