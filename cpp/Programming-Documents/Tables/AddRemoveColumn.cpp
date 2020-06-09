#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

typedef System::SharedPtr<Cell> TCellPtr;

namespace
{
    class Column : public System::Object
    {
        typedef Column ThisType;
        typedef System::Object BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        Column(System::SharedPtr<Table> table, int32_t columnIndex);
        std::vector<TCellPtr> GetColumnCells();
        int32_t IndexOf(System::SharedPtr<Cell> cell);
        System::SharedPtr<Column> InsertColumnBefore();
        void Remove();
        System::String ToTxt();
    private:
        int32_t mColumnIndex;
        System::SharedPtr<Table> mTable;
    };

    RTTI_INFO_IMPL_HASH(3744029511u, Column, ThisTypeBaseTypesInfo);

    Column::Column(System::SharedPtr<Table> table, int32_t columnIndex) : mColumnIndex(0)
    {
        mTable = table;
        mColumnIndex = columnIndex;
    }

    std::vector<TCellPtr> Column::GetColumnCells()
    {
        std::vector<TCellPtr> columnCells;
        for (System::SharedPtr<Row> row : System::IterateOver<System::SharedPtr<Row>>(mTable->get_Rows()))
        {
            TCellPtr cell = row->get_Cells()->idx_get(mColumnIndex);
            if (cell != nullptr)
            {
                columnCells.push_back(cell);
            }
        }
        return columnCells;
    }

    int32_t Column::IndexOf(System::SharedPtr<Cell> cell)
    {
        std::vector<TCellPtr> columnCells(GetColumnCells());
        std::vector<TCellPtr>::iterator iterator = std::find(columnCells.begin(), columnCells.end(), cell);
        return iterator == columnCells.end() ? -1 : std::distance(columnCells.begin(), iterator);
    }

    System::SharedPtr<Column> Column::InsertColumnBefore()
    {
        std::vector<TCellPtr> columnCells(GetColumnCells());
        if (columnCells.empty())
        {
            throw System::ArgumentException(u"Column must not be empty");
        }
        // Create a clone of this column.
        for (System::SharedPtr<Cell> cell : columnCells)
        {
            cell->get_ParentRow()->InsertBefore((System::StaticCast<Node>(cell))->Clone(false), cell);
        }
        // This is the new column.
        System::SharedPtr<Column> column = System::MakeObject<Column>(columnCells[0]->get_ParentRow()->get_ParentTable(), mColumnIndex);
        // We want to make sure that the cells are all valid to work with (have at least one paragraph).
        columnCells = column->GetColumnCells();
        for (System::SharedPtr<Cell> cell : columnCells)
        {
            cell->EnsureMinimum();
        }
        // Increase the index which this column represents since there is now one extra column infront.
        mColumnIndex++;
        return column;
    }

    void Column::Remove()
    {
        for (System::SharedPtr<Cell> cell : GetColumnCells())
        {
            cell->Remove();
        }
    }

    System::String Column::ToTxt()
    {
        System::SharedPtr<System::Text::StringBuilder> builder = System::MakeObject<System::Text::StringBuilder>();
        for (System::SharedPtr<Cell> cell : GetColumnCells())
        {
            builder->Append(cell->ToString(SaveFormat::Text));
        }
        return builder->ToString();
    }

    void RemoveColumn(System::SharedPtr<Document> doc)
    {
        // ExStart:RemoveColumn
        // Get the second table in the document.
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 1, true));
        // Get the third column from the table and remove it.
        System::SharedPtr<Column> column = System::MakeObject<Column>(table, 2);
        column->Remove();
        // ExEnd:RemoveColumn
        std::cout << "Third column removed successfully." << std::endl;
    }

    void InsertBlankColumn(System::SharedPtr<Document> doc)
    {
        // ExStart:InsertBlankColumn
        // Get the first table in the document.
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // Get the second column in the table.
        System::SharedPtr<Column> column = System::MakeObject<Column>(table, 0);
        // Print the plain text of the column to the screen.
        std::cout << column->ToTxt().ToUtf8String() << std::endl;
        // Create a new column to the left of this column.
        // This is the same as using the "Insert Column Before" command in Microsoft Word.
        System::SharedPtr<Column> newColumn = column->InsertColumnBefore();
        // Add some text to each of the column cells.
        std::vector<TCellPtr> columnCells = newColumn->GetColumnCells();
        for (System::SharedPtr<Cell> cell : newColumn->GetColumnCells())
        {
            cell->get_FirstParagraph()->AppendChild(System::MakeObject<Run>(doc, System::String(u"Column Text ") + newColumn->IndexOf(cell)));
        }
        // ExEnd:InsertBlankColumn
        std::cout << "Column added successfully." << std::endl;
    }
}

void AddRemoveColumn()
{
    std::cout << "AddRemoveColumn example started." << std::endl;
    // The path to the documents directory.
    System::String inputFilePath = GetInputDataDir_WorkingWithTables() + u"Table.Document.doc";
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputFilePath);
    InsertBlankColumn(doc);
    RemoveColumn(doc);
    std::cout << "AddRemoveColumn example finished." << std::endl << std::endl;
}