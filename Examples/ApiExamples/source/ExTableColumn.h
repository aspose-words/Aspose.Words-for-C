#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/collections/list.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/string_builder.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace ApiExamples {

class ExTableColumn : public ApiExampleBase
{
public:
    /// <summary>
    /// Represents a facade object for a column of a table in a Microsoft Word document.
    /// </summary>
    class Column : public System::Object
    {
    public:
        /// <summary>
        /// Returns the cells which make up the column.
        /// </summary>
        ArrayPtr<SharedPtr<Cell>> get_Cells()
        {
            return GetColumnCells()->ToArray();
        }

        Column(SharedPtr<Table> table, int columnIndex) : mColumnIndex(0)
        {
            if (table == nullptr)
            {
                throw System::ArgumentException(u"table");
            }

            mTable = table;
            mColumnIndex = columnIndex;
        }

        /// <summary>
        /// Returns a new column facade from the table and supplied zero-based index.
        /// </summary>
        static SharedPtr<ExTableColumn::Column> FromIndex(SharedPtr<Table> table, int columnIndex)
        {
            return MakeObject<ExTableColumn::Column>(table, columnIndex);
        }

        /// <summary>
        /// Returns the index of the given cell in the column.
        /// </summary>
        int IndexOf(SharedPtr<Cell> cell)
        {
            return GetColumnCells()->IndexOf(cell);
        }

        /// <summary>
        /// Inserts a new column before this column into the table.
        /// </summary>
        SharedPtr<ExTableColumn::Column> InsertColumnBefore()
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

            auto newColumn = MakeObject<ExTableColumn::Column>(columnCells[0]->get_ParentRow()->get_ParentTable(), mColumnIndex);

            // We want to make sure that the cells are all valid to work with (have at least one paragraph).
            for (SharedPtr<Cell> cell : newColumn->get_Cells())
            {
                cell->EnsureMinimum();
            }

            // Increment the index of this column represents since there is a new column before it.
            mColumnIndex++;

            return newColumn;
        }

        /// <summary>
        /// Removes the column from the table.
        /// </summary>
        void Remove()
        {
            for (SharedPtr<Cell> cell : get_Cells())
            {
                cell->Remove();
            }
        }

        /// <summary>
        /// Returns the text of the column.
        /// </summary>
        String ToTxt()
        {
            auto builder = MakeObject<System::Text::StringBuilder>();

            for (SharedPtr<Cell> cell : get_Cells())
            {
                builder->Append(cell->ToString(SaveFormat::Text));
            }

            return builder->ToString();
        }

    private:
        int mColumnIndex;
        SharedPtr<Table> mTable;

        /// <summary>
        /// Provides an up-to-date collection of cells which make up the column represented by this facade.
        /// </summary>
        SharedPtr<System::Collections::Generic::List<SharedPtr<Cell>>> GetColumnCells()
        {
            SharedPtr<System::Collections::Generic::List<SharedPtr<Cell>>> columnCells = MakeObject<System::Collections::Generic::List<SharedPtr<Cell>>>();

            for (auto row : System::IterateOver(mTable->get_Rows()->LINQ_OfType<SharedPtr<Row>>()))
            {
                SharedPtr<Cell> cell = row->get_Cells()->idx_get(mColumnIndex);
                if (cell != nullptr)
                {
                    columnCells->Add(cell);
                }
            }

            return columnCells;
        }
    };

    void RemoveColumnFromTable()
    {
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 1, true));

        SharedPtr<ExTableColumn::Column> column = ExTableColumn::Column::FromIndex(table, 2);
        column->Remove();

        doc->Save(ArtifactsDir + u"TableColumn.RemoveColumn.doc");

        ASSERT_EQ(16, table->GetChildNodes(NodeType::Cell, true)->get_Count());
        ASSERT_EQ(u"Cell 7 contents", table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2)->ToString(SaveFormat::Text).Trim());
        ASSERT_EQ(u"Cell 11 contents", table->get_LastRow()->get_Cells()->idx_get(2)->ToString(SaveFormat::Text).Trim());
    }

    void Insert()
    {
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 1, true));

        SharedPtr<ExTableColumn::Column> column = ExTableColumn::Column::FromIndex(table, 1);

        // Create a new column to the left of this column.
        // This is the same as using the "Insert Column Before" command in Microsoft Word.
        SharedPtr<ExTableColumn::Column> newColumn = column->InsertColumnBefore();

        // Add some text to each cell in the column.
        for (SharedPtr<Cell> cell : newColumn->get_Cells())
        {
            cell->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, String(u"Column Text ") + newColumn->IndexOf(cell)));
        }

        doc->Save(ArtifactsDir + u"TableColumn.Insert.doc");

        ASSERT_EQ(24, table->GetChildNodes(NodeType::Cell, true)->get_Count());
        ASSERT_EQ(u"Column Text 0", table->get_FirstRow()->get_Cells()->idx_get(1)->ToString(SaveFormat::Text).Trim());
        ASSERT_EQ(u"Column Text 3", table->get_LastRow()->get_Cells()->idx_get(1)->ToString(SaveFormat::Text).Trim());
    }

    void TableColumnToTxt()
    {
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 1, true));

        SharedPtr<ExTableColumn::Column> column = ExTableColumn::Column::FromIndex(table, 0);
        std::cout << column->ToTxt() << std::endl;

        ASSERT_EQ(u"\r\nRow 1\r\nRow 2\r\nRow 3\r\n", column->ToTxt());
    }
};

} // namespace ApiExamples
