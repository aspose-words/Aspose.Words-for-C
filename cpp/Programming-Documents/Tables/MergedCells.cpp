#include "stdafx.h"
#include "examples.h"

#include <system/shared_ptr.h>
#include <system/special_casts.h>
#include <system/collections/list.h>
#include <system/object.h>
#include <drawing/point.h>
#include <drawing/rectangle.h>

#include "Model/Document/Document.h"
#include "Model/Document/DocumentBuilder.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/TableCollection.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/RowCollection.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Tables/CellMerge.h>
#include <Model/Tables/CellCollection.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/Body.h>
#include <Model/Nodes/NodeType.h>

using namespace Aspose::Words;

namespace
{
    class CellInfo : public System::Object
    {
        typedef CellInfo ThisType;
        typedef System::Object BaseType;

        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        CellInfo(int32_t colSpan, int32_t rowSpan);

        int32_t GetColSpan();
        int32_t GetRowSpan();

    private:
        int32_t mColSpan;
        int32_t mRowSpan;
    };

    RTTI_INFO_IMPL_HASH(1673797985u, CellInfo, ThisTypeBaseTypesInfo);

    CellInfo::CellInfo(int32_t colSpan, int32_t rowSpan) : mColSpan(0), mRowSpan(0)
    {
        mColSpan = colSpan;
        mRowSpan = rowSpan;
    }

    int32_t CellInfo::GetColSpan()
    {
        return mColSpan;
    }

    int32_t CellInfo::GetRowSpan()
    {
        return mRowSpan;
    }

    typedef System::Collections::Generic::List<System::SharedPtr<CellInfo>> TCellsInfo;
    typedef System::SharedPtr<TCellsInfo> TCellsInfoP;

    class RowInfo : public System::Object
    {
        typedef RowInfo ThisType;
        typedef System::Object BaseType;

        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        RowInfo();
        TCellsInfoP GetCells();

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        TCellsInfoP mCells;
    };

    RTTI_INFO_IMPL_HASH(2224295735u, RowInfo, ThisTypeBaseTypesInfo);

    RowInfo::RowInfo() : mCells(System::MakeObject<TCellsInfo>())
    {
    }

    TCellsInfoP RowInfo::GetCells()
    {
        return mCells;
    }

    System::Object::shared_members_type RowInfo::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("RowInfo::mCells", this->mCells);
        return std::move(result);
    }

    typedef System::Collections::Generic::List<System::SharedPtr<RowInfo>> TRowsInfo;
    typedef System::SharedPtr<TRowsInfo> TRowsInfoP;

    class TableInfo : public System::Object
    {
        typedef TableInfo ThisType;
        typedef System::Object BaseType;

        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        TableInfo();
        TRowsInfoP GetRows();

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        TRowsInfoP mRows;
    };

    RTTI_INFO_IMPL_HASH(2175561643u, TableInfo, ThisTypeBaseTypesInfo);

    TableInfo::TableInfo() : mRows(System::MakeObject<TRowsInfo>())
    {
    }

    TRowsInfoP TableInfo::GetRows()
    {
        return mRows;
    }

    System::Object::shared_members_type TableInfo::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("TableInfo::mRows", this->mRows);
        return std::move(result);
    }

    System::String PrintCellMergeType(System::SharedPtr<Aspose::Words::Tables::Cell> cell)
    {
        bool isHorizontallyMerged = cell->get_CellFormat()->get_HorizontalMerge() != Aspose::Words::Tables::CellMerge::None;
        bool isVerticallyMerged = cell->get_CellFormat()->get_VerticalMerge() != Aspose::Words::Tables::CellMerge::None;
        System::String cellLocation = System::String::Format(u"R{0}, C{1}", cell->get_ParentRow()->get_ParentTable()->IndexOf(cell->get_ParentRow()) + 1, cell->get_ParentRow()->IndexOf(cell) + 1);
        if (isHorizontallyMerged && isVerticallyMerged)
        {
            return System::String::Format(u"The cell at {0} is both horizontally and vertically merged", cellLocation);
        }
        if (isHorizontallyMerged)
        {
            return System::String::Format(u"The cell at {0} is horizontally merged.", cellLocation);
        }
        if (isVerticallyMerged)
        {
            return System::String::Format(u"The cell at {0} is vertically merged", cellLocation);
        }
        return System::String::Format(u"The cell at {0} is not merged", cellLocation);
    }

    void CheckCellsMerged(System::String const &dataDir)
    {
        // ExStart:CheckCellsMerged 
        auto doc = System::MakeObject<Document>(dataDir + u"Table.MergedCells.doc");
        // Retrieve the first table in the document.
        auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
        auto row_enumerator = (System::DynamicCastEnumerableTo<System::SharedPtr<Aspose::Words::Tables::Row>>(table->get_Rows()))->GetEnumerator();
        decltype(row_enumerator->get_Current()) row;
        while (row_enumerator->MoveNext() && (row = row_enumerator->get_Current(), true))
        {
            auto cell_enumerator = (System::DynamicCastEnumerableTo<System::SharedPtr<Aspose::Words::Tables::Cell>>(row->get_Cells()))->GetEnumerator();
            decltype(cell_enumerator->get_Current()) cell;
            while (cell_enumerator->MoveNext() && (cell = cell_enumerator->get_Current(), true))
            {
                std::cout << PrintCellMergeType(cell).ToUtf8String() << std::endl;
            }
        }
        // ExEnd:CheckCellsMerged 
    }

    void HorizontalMerge(System::String const &dataDir)
    {
        // ExStart:HorizontalMerge
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        builder->InsertCell();
        builder->get_CellFormat()->set_HorizontalMerge(Aspose::Words::Tables::CellMerge::First);
        builder->Write(u"Text in merged cells.");
        builder->InsertCell();
        // This cell is merged to the previous and should be empty.
        builder->get_CellFormat()->set_HorizontalMerge(Aspose::Words::Tables::CellMerge::Previous);
        builder->EndRow();
        builder->InsertCell();
        builder->get_CellFormat()->set_HorizontalMerge(Aspose::Words::Tables::CellMerge::None);
        builder->Write(u"Text in one cell.");
        builder->InsertCell();
        builder->Write(u"Text in another cell.");
        builder->EndRow();
        builder->EndTable();
        System::String outputPath = dataDir + GetOutputFilePath(u"MergedCells.HorizontalMerge.doc");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:HorizontalMerge
        std::cout << "Table created successfully with cells in the first row horizontally merged." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void VerticalMerge(System::String const &dataDir)
    {
        // ExStart:VerticalMerge
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        builder->InsertCell();
        builder->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::First);
        builder->Write(u"Text in merged cells.");
        builder->InsertCell();
        builder->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::None);
        builder->Write(u"Text in one cell");
        builder->EndRow();
        builder->InsertCell();
        // This cell is vertically merged to the cell above and should be empty.
        builder->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::Previous);
        builder->InsertCell();
        builder->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::None);
        builder->Write(u"Text in another cell");
        builder->EndRow();
        builder->EndTable();
        System::String outputPath = dataDir + GetOutputFilePath(u"MergedCells.VerticalMerge.doc");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:VerticalMerge
        std::cout << "Table created successfully with two columns with cells merged vertically in the first column." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void MergeCells(System::SharedPtr<Aspose::Words::Tables::Cell> startCell, System::SharedPtr<Aspose::Words::Tables::Cell> endCell)
    {
        auto parentTable = startCell->get_ParentRow()->get_ParentTable();
        // Find the row and cell indices for the start and end cell.
        System::Drawing::Point startCellPos(startCell->get_ParentRow()->IndexOf(startCell), parentTable->IndexOf(startCell->get_ParentRow()));
        System::Drawing::Point endCellPos(endCell->get_ParentRow()->IndexOf(endCell), parentTable->IndexOf(endCell->get_ParentRow()));
        // Create the range of cells to be merged based off these indices. Inverse each index if the end cell if before the start cell. 
        System::Drawing::Rectangle mergeRange(System::Math::Min(startCellPos.get_X(), endCellPos.get_X()), System::Math::Min(startCellPos.get_Y(), endCellPos.get_Y()), System::Math::Abs(endCellPos.get_X() - startCellPos.get_X()) + 1, System::Math::Abs(endCellPos.get_Y() - startCellPos.get_Y()) + 1);
        auto row_enumerator = (System::DynamicCastEnumerableTo<System::SharedPtr<Aspose::Words::Tables::Row>>(parentTable->get_Rows()))->GetEnumerator();
        decltype(row_enumerator->get_Current()) row;
        while (row_enumerator->MoveNext() && (row = row_enumerator->get_Current(), true))
        {
            auto cell_enumerator = (System::DynamicCastEnumerableTo<System::SharedPtr<Aspose::Words::Tables::Cell>>(row->get_Cells()))->GetEnumerator();
            decltype(cell_enumerator->get_Current()) cell;
            while (cell_enumerator->MoveNext() && (cell = cell_enumerator->get_Current(), true))
            {
                System::Drawing::Point currentPos(row->IndexOf(cell), parentTable->IndexOf(row));
                // Check if the current cell is inside our merge range then merge it.
                if (mergeRange.Contains(currentPos))
                {
                    if (currentPos.get_X() == mergeRange.get_X())
                    {
                        cell->get_CellFormat()->set_HorizontalMerge(Aspose::Words::Tables::CellMerge::First);
                    }
                    else
                    {
                        cell->get_CellFormat()->set_HorizontalMerge(Aspose::Words::Tables::CellMerge::Previous);
                    }
                    if (currentPos.get_Y() == mergeRange.get_Y())
                    {
                        cell->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::First);
                    }
                    else
                    {
                        cell->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::Previous);
                    }
                }
            }
        }
    }

    void MergeCellRange(System::String const &dataDir)
    {
        // ExStart:MergeCellRange
        // Open the document
        auto doc = System::MakeObject<Document>(dataDir + u"Table.Document.doc");
        // Retrieve the first table in the body of the first section.
        auto table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        // We want to merge the range of cells found inbetween these two cells.
        auto cellStartRange = table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2);
        auto cellEndRange = table->get_Rows()->idx_get(3)->get_Cells()->idx_get(3);
        // Merge all the cells between the two specified cells into one.
        MergeCells(cellStartRange, cellEndRange);
        System::String outputPath = dataDir + GetOutputFilePath(u"MergedCells.MergeCellRange.doc");
        // Save the document.
        doc->Save(outputPath);
        // ExEnd:MergeCellRange
        std::cout << "Cells merged successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void MergedCells()
{
    std::cout << "MergedCells example started." << std::endl;
    // ExStart:MergedCells
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithTables();
    CheckCellsMerged(dataDir);
    // The below method shows how to create a table with two rows with cells in the first row horizontally merged.
    HorizontalMerge(dataDir);
    // The below method shows how to create a table with two columns with cells merged vertically in the first column.
    VerticalMerge(dataDir);
    // The below method shows how to merges the range of cells between the two specified cells.
    MergeCellRange(dataDir);
    // ExEnd:MergedCells
    std::cout << "MergedCells example finished." << std::endl << std::endl;
}