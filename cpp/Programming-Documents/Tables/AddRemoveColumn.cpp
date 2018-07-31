#include "stdafx.h"
#include "examples.h"

#include <system/array.h>
#include <system/shared_ptr.h>
#include <system/special_casts.h>
#include <system/collections/list.h>
#include <system/exceptions.h>
#include <system/object.h>
#include <system/string.h>
#include <system/text/string_builder.h>

#include "Model/Document/Document.h"
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

namespace
{
	// ExStart:ColumnClass
    class Column : public System::Object
    {
        typedef Column ThisType;
        typedef System::Object BaseType;

        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:

        Column(System::SharedPtr<Aspose::Words::Tables::Table> table, int32_t columnIndex);
        System::ArrayPtr<System::SharedPtr<Aspose::Words::Tables::Cell>> GetColumnCells();
        int32_t IndexOf(System::SharedPtr<Aspose::Words::Tables::Cell> cell);
        System::SharedPtr<Column> InsertColumnBefore();
        void Remove();
        System::String ToTxt();

    protected:

        System::Object::shared_members_type GetSharedMembers() override;

    private:

        int32_t mColumnIndex;
        System::SharedPtr<Aspose::Words::Tables::Table> mTable;
    };

    RTTI_INFO_IMPL_HASH(3744029511u, Column, ThisTypeBaseTypesInfo);

    Column::Column(System::SharedPtr<Aspose::Words::Tables::Table> table, int32_t columnIndex) : mColumnIndex(0)
    {
        mTable = table;
        mColumnIndex = columnIndex;
    }

    System::ArrayPtr<System::SharedPtr<Aspose::Words::Tables::Cell>> Column::GetColumnCells()
    {
        auto columnCells = System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Tables::Cell>>>();
        auto row_enumerator = (System::DynamicCastEnumerableTo<System::SharedPtr<Aspose::Words::Tables::Row>>(mTable->get_Rows()))->GetEnumerator();
        decltype(row_enumerator->get_Current()) row;
        while (row_enumerator->MoveNext() && (row = row_enumerator->get_Current(), true))
        {
            auto cell = row->get_Cells()->idx_get(mColumnIndex);
            if (cell != nullptr)
            {
                columnCells->Add(cell);
            }
        }
        return columnCells->ToArray();
    }

    int32_t Column::IndexOf(System::SharedPtr<Aspose::Words::Tables::Cell> cell)
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
        for (int index = 0; index < columnCells->Count(); ++index)
        {
            auto cell = columnCells[index];
            cell->get_ParentRow()->InsertBefore((System::StaticCast<Node>(cell))->Clone(false), cell);
        }
        // This is the new column.
        auto column = System::MakeObject<Column>(columnCells[0]->get_ParentRow()->get_ParentTable(), mColumnIndex);
        // We want to make sure that the cells are all valid to work with (have at least one paragraph).
        columnCells = column->GetColumnCells();
        for (int index = 0; index < columnCells->Count(); ++index)
        {
            auto cell = columnCells[index];
            cell->EnsureMinimum();
        }
        // Increase the index which this column represents since there is now one extra column infront.
        mColumnIndex++;
        return column;
    }

    void Column::Remove()
    {
        for (int index = 0; index < GetColumnCells()->Count(); ++index)
        {
            auto cell = GetColumnCells()[index];
            cell->Remove();
        }
    }

    System::String Column::ToTxt()
    {
        auto builder = System::MakeObject<System::Text::StringBuilder>();
        for (int index = 0; index < GetColumnCells()->Count(); ++index)
        {
            auto cell = GetColumnCells()[index];
            builder->AppendLine(cell->GetText());
        }
        return builder->ToString();
    }

    System::Object::shared_members_type Column::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("Column::mTable", this->mTable);
        return std::move(result);
    }
	// ExEnd:ColumnClass

    void RemoveColumn(System::SharedPtr<Document> doc)
    {
        // ExStart:RemoveColumn
        // Get the second table in the document.
        auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));
        // Get the third column from the table and remove it.
        auto column = System::MakeObject<Column>(table, 2);
        column->Remove();
        // ExEnd:RemoveColumn
        std::cout << "\nThird column removed successfully." << std::endl;
    }

    void InsertBlankColumn(System::SharedPtr<Document> doc)
    {
        // ExStart:InsertBlankColumn
        // Get the first table in the document.
        auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
        // Get the second column in the table.
        auto column = System::MakeObject<Column>(table, 0);
        // Print the plain text of the column to the screen.
        std::cout << column->ToTxt().ToUtf8String() << std::endl;
        // Create a new column to the left of this column.
        // This is the same as using the "Insert Column Before" command in Microsoft Word.
        auto newColumn = column->InsertColumnBefore();
        // Add some text to each of the column cells.
        auto columnCells = newColumn->GetColumnCells();
        for (int index = 0; index < columnCells->Count(); ++index)
        {
            auto cell = columnCells[index];
            cell->get_FirstParagraph()->AppendChild(System::MakeObject<Aspose::Words::Run>(doc, System::String(u"Column Text ") + newColumn->IndexOf(cell)));
        }
        // ExEnd:InsertBlankColumn
        std::cout << "\nColumn added successfully." << std::endl;
    }

}

void AddRemoveColumn()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithTables() + u"Table.Document.doc";
    auto doc = System::MakeObject<Document>(dataDir);
    InsertBlankColumn(doc);
    RemoveColumn(doc);
}