#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <xml/xml_document.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/CellMerge.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;

namespace
{
    // ExStart:HorizontalAndVerticalMergeHelperClasses
    class CellInfo : public System::Object
    {
        typedef CellInfo ThisType;
        typedef System::Object BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        CellInfo(int32_t colSpan, int32_t rowSpan) : mColSpan(colSpan), mRowSpan(rowSpan) {}
        int32_t GetColSpan() { return mColSpan; }
        int32_t GetRowSpan() { return mRowSpan; }

    private:
        int32_t mColSpan;
        int32_t mRowSpan;
    };

    RTTI_INFO_IMPL_HASH(1673797985u, CellInfo, ThisTypeBaseTypesInfo);

    typedef System::SharedPtr<CellInfo> TCellInfoPtr;

    class RowInfo : public System::Object
    {
        typedef RowInfo ThisType;
        typedef System::Object BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        RowInfo()
        {
        }

        std::vector<TCellInfoPtr>& GetCells() { return mCells; }

    private:
        std::vector<TCellInfoPtr> mCells;
    };

    RTTI_INFO_IMPL_HASH(2224295735u, RowInfo, ThisTypeBaseTypesInfo);

    typedef System::SharedPtr<RowInfo> TRowInfoPtr;

    class TableInfo : public System::Object
    {
        typedef TableInfo ThisType;
        typedef System::Object BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        TableInfo()
        {
        }
        std::vector<TRowInfoPtr>& GetRows() { return mRows; }

    private:
        std::vector<TRowInfoPtr> mRows;
    };

    RTTI_INFO_IMPL_HASH(2175561643u, TableInfo, ThisTypeBaseTypesInfo);

    typedef System::SharedPtr<TableInfo> TTableInfoPtr;

    // ExEnd:HorizontalAndVerticalMergeHelperClasses

    class SpanVisitor : public DocumentVisitor
    {
        typedef SpanVisitor ThisType;
        typedef DocumentVisitor BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        SpanVisitor(System::SharedPtr<Document> doc);
        virtual VisitorAction VisitCellStart(System::SharedPtr<Cell> cell);
    private:
        std::vector<TTableInfoPtr> mTables;
        System::SharedPtr<NodeCollection> mWordTables;
    };

    RTTI_INFO_IMPL_HASH(4097796275u, SpanVisitor, ThisTypeBaseTypesInfo);

    SpanVisitor::SpanVisitor(System::SharedPtr<Document> doc) : mWordTables(nullptr)
    {
        // Self reference protector
        System::Details::ThisProtector __local_self_ref(this);

        mWordTables = doc->GetChildNodes(NodeType::Table, true);
        System::SharedPtr<System::IO::MemoryStream> htmlStream = System::MakeObject<System::IO::MemoryStream>();
        System::SharedPtr<HtmlSaveOptions> options = System::MakeObject<HtmlSaveOptions>();
        options->set_ImagesFolder(System::IO::Path::GetTempPath());
        doc->Save(htmlStream, options);
        System::SharedPtr<System::Xml::XmlDocument> xmlDoc = System::MakeObject<System::Xml::XmlDocument>();
        htmlStream->set_Position(0);
        xmlDoc->Load(htmlStream);
        System::SharedPtr<System::Xml::XmlNodeList> tables = xmlDoc->get_DocumentElement()->SelectNodes(u"// Table");
        for (System::SharedPtr<System::Xml::XmlNode> table : System::IterateOver(tables))
        {
            TTableInfoPtr tableInf = System::MakeObject<TableInfo>();
            // Get collection of rows in the table
            System::SharedPtr<System::Xml::XmlNodeList> rows = table->SelectNodes(u"tr");

            for (System::SharedPtr<System::Xml::XmlNode> row : System::IterateOver(rows))
            {
                TRowInfoPtr rowInf = System::MakeObject<RowInfo>();

                // Get collection of cells
                System::SharedPtr<System::Xml::XmlNodeList> cells = row->SelectNodes(u"td");

                for (System::SharedPtr<System::Xml::XmlNode> cell : System::IterateOver(cells))
                {
                    // Determine row span and colspan of the current cell
                    System::SharedPtr<System::Xml::XmlAttribute> colSpanAttr = cell->get_Attributes()->idx_get(u"colspan");
                    System::SharedPtr<System::Xml::XmlAttribute> rowSpanAttr = cell->get_Attributes()->idx_get(u"rowspan");

                    int32_t colSpan = colSpanAttr == nullptr ? 0 : System::Convert::ToInt32(colSpanAttr->get_Value());
                    int32_t rowSpan = rowSpanAttr == nullptr ? 0 : System::Convert::ToInt32(rowSpanAttr->get_Value());

                    TCellInfoPtr cellInf = System::MakeObject<CellInfo>(colSpan, rowSpan);
                    rowInf->GetCells().push_back(cellInf);
                }
                tableInf->GetRows().push_back(rowInf);
            }
            mTables.push_back(tableInf);
        }
    }

    VisitorAction SpanVisitor::VisitCellStart(System::SharedPtr<Cell> cell)
    {
        // Determone index of current table
        int32_t tabIdx = mWordTables->IndexOf(cell->get_ParentRow()->get_ParentTable());

        // Determine index of current row
        int32_t rowIdx = cell->get_ParentRow()->get_ParentTable()->IndexOf(cell->get_ParentRow());

        // And determine index of current cell
        int32_t cellIdx = cell->get_ParentRow()->IndexOf(cell);

        // Determine colspan and rowspan of current cell
        int32_t colSpan = 0;
        int32_t rowSpan = 0;
        if (tabIdx < mTables.size() && rowIdx < mTables[tabIdx]->GetRows().size() &&
            cellIdx < mTables[tabIdx]->GetRows()[rowIdx]->GetCells().size())
        {
            colSpan = mTables[tabIdx]->GetRows()[rowIdx]->GetCells()[cellIdx]->GetColSpan();
            rowSpan = mTables[tabIdx]->GetRows()[rowIdx]->GetCells()[cellIdx]->GetRowSpan();
        }

        std::cout << tabIdx << "." << rowIdx << "." << cellIdx << "colspan=" << colSpan << "\t rowspan=" << rowSpan << std::endl;
        return VisitorAction::Continue;
    }

    // ExStart:PrintCellMergeType
    System::String PrintCellMergeType(System::SharedPtr<Cell> cell)
    {
        bool isHorizontallyMerged = cell->get_CellFormat()->get_HorizontalMerge() != CellMerge::None;
        bool isVerticallyMerged = cell->get_CellFormat()->get_VerticalMerge() != CellMerge::None;
        System::String cellLocation = System::String::Format(u"R{0}, C{1}", cell->get_ParentRow()->get_ParentTable()->IndexOf(cell->get_ParentRow()) + 1, cell->get_ParentRow()->IndexOf(cell) + 1);
        if (isHorizontallyMerged && isVerticallyMerged)
        {
            return System::String::Format(u"The cell at {0} is both horizontally and vertically merged", cellLocation);
        }
        else if (isHorizontallyMerged)
        {
            return System::String::Format(u"The cell at {0} is horizontally merged.", cellLocation);
        }
        else if (isVerticallyMerged)
        {
            return System::String::Format(u"The cell at {0} is vertically merged", cellLocation);
        }
        return System::String::Format(u"The cell at {0} is not merged", cellLocation);
    }
    // ExEnd:PrintCellMergeType

    void CheckCellsMerged(System::String const &inputDataDir)
    {
        // ExStart:CheckCellsMerged 
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.MergedCells.doc");
        // Retrieve the first table in the document.
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        for (System::SharedPtr<Row> row : System::IterateOver<System::SharedPtr<Row>>(table->get_Rows()))
        {
            for (System::SharedPtr<Cell> cell : System::IterateOver<System::SharedPtr<Cell>>(row->get_Cells()))
            {
                std::cout << PrintCellMergeType(cell).ToUtf8String() << std::endl;
            }
        }
        // ExEnd:CheckCellsMerged 
    }

    void HorizontalMerge(System::String const &outputDataDir)
    {
        // ExStart:HorizontalMerge
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        builder->InsertCell();
        builder->get_CellFormat()->set_HorizontalMerge(CellMerge::First);
        builder->Write(u"Text in merged cells.");
        builder->InsertCell();
        // This cell is merged to the previous and should be empty.
        builder->get_CellFormat()->set_HorizontalMerge(CellMerge::Previous);
        builder->EndRow();
        builder->InsertCell();
        builder->get_CellFormat()->set_HorizontalMerge(CellMerge::None);
        builder->Write(u"Text in one cell.");
        builder->InsertCell();
        builder->Write(u"Text in another cell.");
        builder->EndRow();
        builder->EndTable();
        System::String outputPath = outputDataDir + u"MergedCells.HorizontalMerge.doc";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:HorizontalMerge
        std::cout << "Table created successfully with cells in the first row horizontally merged." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void VerticalMerge(System::String const &outputDataDir)
    {
        // ExStart:VerticalMerge
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        builder->InsertCell();
        builder->get_CellFormat()->set_VerticalMerge(CellMerge::First);
        builder->Write(u"Text in merged cells.");
        builder->InsertCell();
        builder->get_CellFormat()->set_VerticalMerge(CellMerge::None);
        builder->Write(u"Text in one cell");
        builder->EndRow();
        builder->InsertCell();
        // This cell is vertically merged to the cell above and should be empty.
        builder->get_CellFormat()->set_VerticalMerge(CellMerge::Previous);
        builder->InsertCell();
        builder->get_CellFormat()->set_VerticalMerge(CellMerge::None);
        builder->Write(u"Text in another cell");
        builder->EndRow();
        builder->EndTable();
        System::String outputPath = outputDataDir + u"MergedCells.VerticalMerge.doc";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:VerticalMerge
        std::cout << "Table created successfully with two columns with cells merged vertically in the first column." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    // ExStart:MergeCells
    void MergeCells(System::SharedPtr<Cell> startCell, System::SharedPtr<Cell> endCell)
    {
        System::SharedPtr<Table> parentTable = startCell->get_ParentRow()->get_ParentTable();
        // Find the row and cell indices for the start and end cell.
        System::Drawing::Point startCellPos(startCell->get_ParentRow()->IndexOf(startCell), parentTable->IndexOf(startCell->get_ParentRow()));
        System::Drawing::Point endCellPos(endCell->get_ParentRow()->IndexOf(endCell), parentTable->IndexOf(endCell->get_ParentRow()));
        // Create the range of cells to be merged based off these indices. Inverse each index if the end cell if before the start cell.
        System::Drawing::Rectangle mergeRange(System::Math::Min(startCellPos.get_X(), endCellPos.get_X()),
                                              System::Math::Min(startCellPos.get_Y(), endCellPos.get_Y()),
                                              System::Math::Abs(endCellPos.get_X() - startCellPos.get_X()) + 1,
                                              System::Math::Abs(endCellPos.get_Y() - startCellPos.get_Y()) + 1);
        for (System::SharedPtr<Row> row : System::IterateOver<System::SharedPtr<Row>>(parentTable->get_Rows()))
        {
            for (System::SharedPtr<Cell> cell : System::IterateOver<System::SharedPtr<Cell>>(row->get_Cells()))
            {
                System::Drawing::Point currentPos(row->IndexOf(cell), parentTable->IndexOf(row));
                // Check if the current cell is inside our merge range then merge it.
                if (mergeRange.Contains(currentPos))
                {
                    if (currentPos.get_X() == mergeRange.get_X())
                    {
                        cell->get_CellFormat()->set_HorizontalMerge(CellMerge::First);
                    }
                    else
                    {
                        cell->get_CellFormat()->set_HorizontalMerge(CellMerge::Previous);
                    }
                    if (currentPos.get_Y() == mergeRange.get_Y())
                    {
                        cell->get_CellFormat()->set_VerticalMerge(CellMerge::First);
                    }
                    else
                    {
                        cell->get_CellFormat()->set_VerticalMerge(CellMerge::Previous);
                    }
                }
            }
        }
    }
    // ExEnd:MergeCells

    void MergeCellRange(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:MergeCellRange
        // Open the document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.Document.doc");
        // Retrieve the first table in the body of the first section.
        System::SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        // We want to merge the range of cells found in between these two cells.
        System::SharedPtr<Cell> cellStartRange = table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2);
        System::SharedPtr<Cell> cellEndRange = table->get_Rows()->idx_get(3)->get_Cells()->idx_get(3);
        // Merge all the cells between the two specified cells into one.
        MergeCells(cellStartRange, cellEndRange);
        System::String outputPath = outputDataDir + u"MergedCells.MergeCellRange.doc";
        // Save the document.
        doc->Save(outputPath);
        // ExEnd:MergeCellRange
        std::cout << "Cells merged successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void PrintHorizontalAndVerticalMerged(System::String const &inputDataDir)
    {
        // ExStart:PrintHorizontalAndVerticalMerged
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.MergedCells.doc");

        // Create visitor
        System::SharedPtr<SpanVisitor> visitor = System::MakeObject<SpanVisitor>(doc);

        // Accept visitor
        doc->Accept(visitor);
        // ExEnd:PrintHorizontalAndVerticalMerged
        std::cout << "Horizontal and vertical merged of a cell prints successfully." << std::endl;
    }

	void ConvertToHorizontallyMergedCells(System::String const &inputDataDir, System::String const &outputDataDir)
	{
		// ExStart:ConvertToHorizontallyMergedCells         
		System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TableXXX.doc");

        auto table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
		table->ConvertToHorizontallyMergedCells();   // Now merged cells have appropriate merge flags.

        System::String outputPath = outputDataDir + u"HorizontallyMergedCells.doc";
        // Save the document.
        doc->Save(outputPath);
		// ExEnd:ConvertToHorizontallyMergedCells

        std::cout << "Now merged cells have appropriate merge flags.\nFile saved at " << outputPath.ToUtf8String() << '\n';
	}
}

void MergedCells()
{
    std::cout << "MergedCells example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithTables();
    System::String outputDataDir = GetOutputDataDir_WorkingWithTables();
    CheckCellsMerged(inputDataDir);
    // The below method shows how to create a table with two rows with cells in the first row horizontally merged.
    HorizontalMerge(outputDataDir);
    // The below method shows how to create a table with two columns with cells merged vertically in the first column.
    VerticalMerge(outputDataDir);
    // The below method shows how to merges the range of cells between the two specified cells.
    MergeCellRange(inputDataDir, outputDataDir);
    // Show how to prints the horizontal and vertical merge of a cell.
    PrintHorizontalAndVerticalMerged(inputDataDir);
    // This method converts cells which are horizontally merged by its width to the cell horizontally merged by flags
#if 0
    // Source document is missing
    ConvertToHorizontallyMergedCells(inputDataDir, outputDataDir); 
#endif
    std::cout << "MergedCells example finished." << std::endl << std::endl;
}