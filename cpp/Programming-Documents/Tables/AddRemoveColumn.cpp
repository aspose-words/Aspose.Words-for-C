#include "stdafx.h"
#include "examples.h"

#include <system/array.h>
#include <system/enumerator_adapter.h>
#include <system/shared_ptr.h>
#include <system/special_casts.h>
#include <system/collections/list.h>
#include <system/exceptions.h>
#include <system/object.h>
#include <system/string.h>
#include <system/text/string_builder.h>
#include "Model/Document/Document.h"
#include <Model/Document/SaveFormat.h>
#include <Model/Text/Run.h>
#include <Model/Text/Paragraph.h>
#include <Model/Tables/Table.h>
#include <Model/Tables/RowCollection.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/CellCollection.h>
#include <Model/Tables/Cell.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/NodeType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

typedef System::Collections::Generic::List<System::SharedPtr<Cell>> TCellList;
typedef System::ArrayPtr<System::SharedPtr<Cell>> TCellArrayPtr;

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
        TCellArrayPtr GetColumnCells();
        int32_t IndexOf(System::SharedPtr<Cell> cell);
        System::SharedPtr<Column> InsertColumnBefore();
        void Remove();
        System::String ToTxt();

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

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

    TCellArrayPtr Column::GetColumnCells()
    {
        auto columnCells = System::MakeObject<TCellList>();
        for (System::SharedPtr<Row> row : System::IterateOver<System::SharedPtr<Row>>(mTable->get_Rows()))
        {
            auto cell = row->get_Cells()->idx_get(mColumnIndex);
            if (cell != nullptr)
            {
                columnCells->Add(cell);
            }
        }
        return columnCells->ToArray();
    }

    int32_t Column::IndexOf(System::SharedPtr<Cell> cell)
    {
        return GetColumnCells()->IndexOf(cell);
    }

    System::SharedPtr<Column> Column::InsertColumnBefore()
    {
        auto columnCells = GetColumnCells();
        if (columnCells->get_Length() == 0)
        {
            throw System::ArgumentException(u"Column must not be empty");
        }
        // Create a clone of this column.
        for (System::SharedPtr<Cell> cell : columnCells)
        {
            cell->get_ParentRow()->InsertBefore((System::StaticCast<Node>(cell))->Clone(false), cell);
        }
        // This is the new column.
        auto column = System::MakeObject<Column>(columnCells[0]->get_ParentRow()->get_ParentTable(), mColumnIndex);
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
        auto builder = System::MakeObject<System::Text::StringBuilder>();
        for (System::SharedPtr<Cell> cell : GetColumnCells())
        {
            builder->Append(cell->ToString(SaveFormat::Text));
        }
        return builder->ToString();
    }

    System::Object::shared_members_type Column::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("Column::mTable", this->mTable);
        return result;
    }

    void RemoveColumn(System::SharedPtr<Document> doc)
    {
        // ExStart:RemoveColumn
        // Get the second table in the document.
        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 1, true));
        // Get the third column from the table and remove it.
        auto column = System::MakeObject<Column>(table, 2);
        column->Remove();
        // ExEnd:RemoveColumn
        std::cout << "Third column removed successfully." << std::endl;
    }

    void InsertBlankColumn(System::SharedPtr<Document> doc)
    {
        // ExStart:InsertBlankColumn
        // Get the first table in the document.
        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // Get the second column in the table.
        auto column = System::MakeObject<Column>(table, 0);
        // Print the plain text of the column to the screen.
        std::cout << column->ToTxt().ToUtf8String() << std::endl;
        // Create a new column to the left of this column.
        // This is the same as using the "Insert Column Before" command in Microsoft Word.
        auto newColumn = column->InsertColumnBefore();
        // Add some text to each of the column cells.
        auto columnCells = newColumn->GetColumnCells();
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
    System::String dataDir = GetDataDir_WorkingWithTables() + u"Table.Document.doc";
    auto doc = System::MakeObject<Document>(dataDir);
    InsertBlankColumn(doc);
    RemoveColumn(doc);
    std::cout << "AddRemoveColumn example finished." << std::endl << std::endl;
}