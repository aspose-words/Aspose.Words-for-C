// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTableColumn.h"

#include <testing/test_predicates.h>
#include <system/text/string_builder.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/console.h>
#include <system/collections/list.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2513584452u, ::ApiExamples::ExTableColumn::Column, ThisTypeBaseTypesInfo);

ArrayPtr<SharedPtr<Aspose::Words::Tables::Cell>> ExTableColumn::Column::get_Cells()
{
    return GetColumnCells()->ToArray();
}

ExTableColumn::Column::Column(SharedPtr<Aspose::Words::Tables::Table> table, int columnIndex)
     : mColumnIndex(0)
{
    if (table == nullptr)
    {
        throw System::ArgumentException(u"table");
    }

    mTable = table;
    mColumnIndex = columnIndex;
}

SharedPtr<ExTableColumn::Column> ExTableColumn::Column::FromIndex(SharedPtr<Aspose::Words::Tables::Table> table, int columnIndex)
{
    return MakeObject<ExTableColumn::Column>(table, columnIndex);
}

int ExTableColumn::Column::IndexOf(SharedPtr<Aspose::Words::Tables::Cell> cell)
{
    return GetColumnCells()->IndexOf(cell);
}

SharedPtr<ExTableColumn::Column> ExTableColumn::Column::InsertColumnBefore()
{
    ArrayPtr<SharedPtr<Cell>> columnCells = get_Cells();

    if (columnCells->get_Length() == 0)
    {
        throw System::ArgumentException(u"Column must not be empty");
    }

    // Create a clone of this column
    for (SharedPtr<Cell> cell : columnCells)
    {
        cell->get_ParentRow()->InsertBefore(cell->Clone(false), cell);
    }

    // This is the new column
    auto column = MakeObject<ExTableColumn::Column>(columnCells[0]->get_ParentRow()->get_ParentTable(), mColumnIndex);

    // We want to make sure that the cells are all valid to work with (have at least one paragraph)
    for (SharedPtr<Cell> cell : column->get_Cells())
    {
        cell->EnsureMinimum();
    }

    // Increase the index which this column represents since there is now one extra column infront
    mColumnIndex++;

    return column;
}

void ExTableColumn::Column::Remove()
{
    for (SharedPtr<Cell> cell : get_Cells())
    {
        cell->Remove();
    }

}

String ExTableColumn::Column::ToTxt()
{
    auto builder = MakeObject<System::Text::StringBuilder>();

    for (SharedPtr<Cell> cell : get_Cells())
    {
        builder->Append(cell->ToString(Aspose::Words::SaveFormat::Text));
    }

    return builder->ToString();
}

SharedPtr<System::Collections::Generic::List<SharedPtr<Aspose::Words::Tables::Cell>>> ExTableColumn::Column::GetColumnCells()
{
    SharedPtr<System::Collections::Generic::List<SharedPtr<Cell>>> columnCells = MakeObject<System::Collections::Generic::List<SharedPtr<Cell>>>();

    for (auto row : System::IterateOver(mTable->get_Rows()->LINQ_OfType<SharedPtr<Row> >()))
    {
        SharedPtr<Cell> cell = row->get_Cells()->idx_get(mColumnIndex);
        if (cell != nullptr)
        {
            columnCells->Add(cell);
        }
    }

    return columnCells;
}

System::Object::shared_members_type ApiExamples::ExTableColumn::Column::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExTableColumn::Column::mTable", this->mTable);

    return result;
}

namespace gtest_test
{

class ExTableColumn : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExTableColumn> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExTableColumn>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExTableColumn> ExTableColumn::s_instance;

} // namespace gtest_test

void ExTableColumn::RemoveColumnFromTable()
{
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));

    // Get the third column from the table and remove it
    SharedPtr<ExTableColumn::Column> column = ExTableColumn::Column::FromIndex(table, 2);
    column->Remove();

    doc->Save(ArtifactsDir + u"TableColumn.RemoveColumn.doc");

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
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));

    // Get the second column in the table
    SharedPtr<ExTableColumn::Column> column = ExTableColumn::Column::FromIndex(table, 1);

    // Create a new column to the left of this column
    // This is the same as using the "Insert Column Before" command in Microsoft Word
    SharedPtr<ExTableColumn::Column> newColumn = column->InsertColumnBefore();

    // Add some text to each of the column cells
    for (SharedPtr<Cell> cell : newColumn->get_Cells())
    {
        cell->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, String(u"Column Text ") + newColumn->IndexOf(cell)));
    }

    doc->Save(ArtifactsDir + u"TableColumn.Insert.doc");

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
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));

    // Get the first column in the table
    SharedPtr<ExTableColumn::Column> column = ExTableColumn::Column::FromIndex(table, 0);

    // Print the plain text of the column to the screen
    System::Console::WriteLine(column->ToTxt());

    ASSERT_EQ(u"\r\nRow 1\r\nRow 2\r\nRow 3\r\n", column->ToTxt());
}

namespace gtest_test
{

TEST_F(ExTableColumn, TableColumnToTxt)
{
    s_instance->TableColumnToTxt();
}

} // namespace gtest_test

} // namespace ApiExamples
