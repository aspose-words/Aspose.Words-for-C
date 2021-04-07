#pragma once
// CPPDEFECT: System.Data is not supported

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Borders/Shading.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Tables/AutoFitBehavior.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/CellMerge.h>
#include <Aspose.Words.Cpp/Model/Tables/CellVerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidthType.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/TableAlignment.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/TextWrapping.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/HeightRule.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <drawing/color.h>
#include <drawing/point.h>
#include <drawing/rectangle.h>
#include <system/array.h>
#include <system/collections/list.h>
#include <system/convert.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/io/memory_stream.h>
#include <system/io/path.h>
#include <system/math.h>
#include <system/primitive_types.h>
#include <system/scope_guard.h>
#include <system/text/string_builder.h>
#include <xml/xml_attribute.h>
#include <xml/xml_attribute_collection.h>
#include <xml/xml_document.h>
#include <xml/xml_element.h>
#include <xml/xml_node.h>
#include <xml/xml_node_list.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Tables {

class WorkingWithTables : public DocsExamplesBase
{
public:
    class RowInfo;
    class CellInfo;

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

        /// <summary>
        /// Returns a new column facade from the table and supplied zero-based index.
        /// </summary>
        static SharedPtr<WorkingWithTables::Column> FromIndex(SharedPtr<Table> table, int columnIndex)
        {
            return WorkingWithTables::Column::MakeObject(table, columnIndex);
        }

        /// <summary>
        /// Returns the index of the given cell in the column.
        /// </summary>
        int IndexOf(SharedPtr<Cell> cell)
        {
            return GetColumnCells()->IndexOf(cell);
        }

        /// <summary>
        /// Inserts a brand new column before this column into the table.
        /// </summary>
        SharedPtr<WorkingWithTables::Column> InsertColumnBefore()
        {
            ArrayPtr<SharedPtr<Cell>> columnCells = get_Cells();

            if (columnCells->get_Length() == 0)
            {
                throw System::ArgumentException(u"Column must not be empty");
            }

            // Create a clone of this column.
            for (SharedPtr<Cell> cell : columnCells)
            {
                cell->get_ParentRow()->InsertBefore(cell->Clone(false), cell);
            }

            // This is the new column.
            auto column = WorkingWithTables::Column::MakeObject(columnCells[0]->get_ParentRow()->get_ParentTable(), mColumnIndex);

            // We want to make sure that the cells are all valid to work with (have at least one paragraph).
            for (SharedPtr<Cell> cell : column->get_Cells())
            {
                cell->EnsureMinimum();
            }

            // Increase the index which this column represents since there is now one extra column in front.
            mColumnIndex++;

            return column;
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
            auto builder = System::MakeObject<System::Text::StringBuilder>();

            for (SharedPtr<Cell> cell : get_Cells())
            {
                builder->Append(cell->ToString(SaveFormat::Text));
            }

            return builder->ToString();
        }

    private:
        int mColumnIndex;
        SharedPtr<Table> mTable;

        Column(SharedPtr<Table> table, int columnIndex) : mColumnIndex(0)
        {
            if (table == nullptr)
            {
                throw System::ArgumentException(u"table");
            }

            mTable = table;
            mColumnIndex = columnIndex;
        }

        MEMBER_FUNCTION_MAKE_OBJECT(Column, CODEPORTING_ARGS(SharedPtr<Table> table, int columnIndex), CODEPORTING_ARGS(table, columnIndex));
        /// <summary>
        /// Provides an up-to-date collection of cells which make up the column represented by this facade.
        /// </summary>
        SharedPtr<System::Collections::Generic::List<SharedPtr<Cell>>> GetColumnCells()
        {
            SharedPtr<System::Collections::Generic::List<SharedPtr<Cell>>> columnCells =
                System::MakeObject<System::Collections::Generic::List<SharedPtr<Cell>>>();

            for (auto row : System::IterateOver<Row>(mTable->get_Rows()))
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

    /// <summary>
    /// Helper class that contains collection of rowinfo for each row.
    /// </summary>
    class TableInfo : public System::Object
    {
    public:
        SharedPtr<System::Collections::Generic::List<SharedPtr<WorkingWithTables::RowInfo>>> get_Rows()
        {
            return mRows;
        }

        TableInfo() : mRows(MakeObject<System::Collections::Generic::List<SharedPtr<WorkingWithTables::RowInfo>>>())
        {
        }

    private:
        SharedPtr<System::Collections::Generic::List<SharedPtr<WorkingWithTables::RowInfo>>> mRows;
    };

    /// <summary>
    /// Helper class that contains collection of cellinfo for each cell.
    /// </summary>
    class RowInfo : public System::Object
    {
    public:
        SharedPtr<System::Collections::Generic::List<SharedPtr<WorkingWithTables::CellInfo>>> get_Cells()
        {
            return mCells;
        }

        RowInfo() : mCells(MakeObject<System::Collections::Generic::List<SharedPtr<WorkingWithTables::CellInfo>>>())
        {
        }

    private:
        SharedPtr<System::Collections::Generic::List<SharedPtr<WorkingWithTables::CellInfo>>> mCells;
    };

    /// <summary>
    /// Helper class that contains info about cell. currently here is only colspan and rowspan.
    /// </summary>
    class CellInfo : public System::Object
    {
    public:
        int get_ColSpan()
        {
            return pr_ColSpan;
        }

        int get_RowSpan()
        {
            return pr_RowSpan;
        }

        CellInfo(int colSpan, int rowSpan) : pr_ColSpan(0), pr_RowSpan(0)
        {
            // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            System::Details::ThisProtector __local_self_ref(this);
            //---------------------------------------------------------Self Reference

            pr_ColSpan = colSpan;
            pr_RowSpan = rowSpan;
        }

    private:
        int pr_ColSpan;
        int pr_RowSpan;
    };

    class SpanVisitor : public DocumentVisitor
    {
    public:
        /// <summary>
        /// Creates new SpanVisitor instance.
        /// </summary>
        /// <param name="doc">
        /// Is document which we should parse.
        /// </param>
        SpanVisitor(SharedPtr<Document> doc) : mTables(MakeObject<System::Collections::Generic::List<SharedPtr<WorkingWithTables::TableInfo>>>())
        {
            // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            System::Details::ThisProtector __local_self_ref(this);
            //---------------------------------------------------------Self Reference

            mWordTables = doc->GetChildNodes(NodeType::Table, true);

            // We will parse HTML to determine the rowspan and colspan of each cell.
            auto htmlStream = MakeObject<System::IO::MemoryStream>();

            auto options = MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
            options->set_ImagesFolder(System::IO::Path::GetTempPath());

            doc->Save(htmlStream, options);

            // Load HTML into the XML document.
            auto xmlDoc = MakeObject<System::Xml::XmlDocument>();
            htmlStream->set_Position(0);
            xmlDoc->Load(htmlStream);

            // Get collection of tables in the HTML document.
            SharedPtr<System::Xml::XmlNodeList> tables = xmlDoc->get_DocumentElement()->SelectNodes(u"// Table");

            for (auto table : System::IterateOver(tables))
            {
                auto tableInf = MakeObject<WorkingWithTables::TableInfo>();
                // Get collection of rows in the table.
                SharedPtr<System::Xml::XmlNodeList> rows = table->SelectNodes(u"tr");

                for (auto row : System::IterateOver(rows))
                {
                    auto rowInf = MakeObject<WorkingWithTables::RowInfo>();
                    // Get collection of cells.
                    SharedPtr<System::Xml::XmlNodeList> cells = row->SelectNodes(u"td");

                    for (auto cell : System::IterateOver(cells))
                    {
                        // Determine row span and colspan of the current cell.
                        SharedPtr<System::Xml::XmlAttribute> colSpanAttr = cell->get_Attributes()->idx_get(u"colspan");
                        SharedPtr<System::Xml::XmlAttribute> rowSpanAttr = cell->get_Attributes()->idx_get(u"rowspan");

                        int colSpan = colSpanAttr == nullptr ? 0 : System::Int32::Parse(colSpanAttr->get_Value());
                        int rowSpan = rowSpanAttr == nullptr ? 0 : System::Int32::Parse(rowSpanAttr->get_Value());

                        auto cellInf = MakeObject<WorkingWithTables::CellInfo>(colSpan, rowSpan);
                        rowInf->get_Cells()->Add(cellInf);
                    }

                    tableInf->get_Rows()->Add(rowInf);
                }

                mTables->Add(tableInf);
            }
        }

        VisitorAction VisitCellStart(SharedPtr<Cell> cell) override
        {
            int tabIdx = mWordTables->IndexOf(cell->get_ParentRow()->get_ParentTable());
            int rowIdx = cell->get_ParentRow()->get_ParentTable()->IndexOf(cell->get_ParentRow());
            int cellIdx = cell->get_ParentRow()->IndexOf(cell);

            int colSpan = 0;
            int rowSpan = 0;
            if (tabIdx < mTables->get_Count() && rowIdx < mTables->idx_get(tabIdx)->get_Rows()->get_Count() &&
                cellIdx < mTables->idx_get(tabIdx)->get_Rows()->idx_get(rowIdx)->get_Cells()->get_Count())
            {
                colSpan = mTables->idx_get(tabIdx)->get_Rows()->idx_get(rowIdx)->get_Cells()->idx_get(cellIdx)->get_ColSpan();
                rowSpan = mTables->idx_get(tabIdx)->get_Rows()->idx_get(rowIdx)->get_Cells()->idx_get(cellIdx)->get_RowSpan();
            }

            std::cout << tabIdx << "." << rowIdx << "." << cellIdx << " colspan=" << colSpan << "\t rowspan=" << rowSpan << std::endl;

            return VisitorAction::Continue;
        }

    private:
        SharedPtr<System::Collections::Generic::List<SharedPtr<WorkingWithTables::TableInfo>>> mTables;
        SharedPtr<NodeCollection> mWordTables;
    };

public:
    void RemoveColumn()
    {
        //ExStart:RemoveColumn
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 1, true));

        SharedPtr<WorkingWithTables::Column> column = WorkingWithTables::Column::FromIndex(table, 2);
        column->Remove();
        //ExEnd:RemoveColumn
    }

    void InsertBlankColumn()
    {
        //ExStart:InsertBlankColumn
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        //ExStart:GetPlainText
        SharedPtr<WorkingWithTables::Column> column = WorkingWithTables::Column::FromIndex(table, 0);
        // Print the plain text of the column to the screen.
        std::cout << column->ToTxt() << std::endl;
        //ExEnd:GetPlainText

        // Create a new column to the left of this column.
        // This is the same as using the "Insert Column Before" command in Microsoft Word.
        SharedPtr<WorkingWithTables::Column> newColumn = column->InsertColumnBefore();

        for (SharedPtr<Cell> cell : newColumn->get_Cells())
        {
            cell->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, String(u"Column Text ") + newColumn->IndexOf(cell)));
        }

        //ExEnd:InsertBlankColumn
    }

    void AutoFitTableToContents()
    {
        //ExStart:AutoFitTableToContents
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        table->AutoFit(AutoFitBehavior::AutoFitToContents);

        doc->Save(ArtifactsDir + u"WorkingWithTables.AutoFitTableToContents.docx");
        //ExEnd:AutoFitTableToContents
    }

    void AutoFitTableToFixedColumnWidths()
    {
        //ExStart:AutoFitTableToFixedColumnWidths
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // Disable autofitting on this table.
        table->AutoFit(AutoFitBehavior::FixedColumnWidths);

        doc->Save(ArtifactsDir + u"WorkingWithTables.AutoFitTableToFixedColumnWidths.docx");
        //ExEnd:AutoFitTableToFixedColumnWidths
    }

    void AutoFitTableToPageWidth()
    {
        //ExStart:AutoFitTableToPageWidth
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // Autofit the first table to the page width.
        table->AutoFit(AutoFitBehavior::AutoFitToWindow);

        doc->Save(ArtifactsDir + u"WorkingWithTables.AutoFitTableToWindow.docx");
        //ExEnd:AutoFitTableToPageWidth
    }

    void CloneCompleteTable()
    {
        //ExStart:CloneCompleteTable
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        // Clone the table and insert it into the document after the original.
        auto tableClone = System::DynamicCast<Table>(table->Clone(true));
        table->get_ParentNode()->InsertAfter(tableClone, table);

        // Insert an empty paragraph between the two tables,
        // or else they will be combined into one upon saving this has to do with document validation.
        table->get_ParentNode()->InsertAfter(MakeObject<Paragraph>(doc), table);

        doc->Save(ArtifactsDir + u"WorkingWithTables.CloneCompleteTable.docx");
        //ExEnd:CloneCompleteTable
    }

    void CloneLastRow()
    {
        //ExStart:CloneLastRow
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        auto clonedRow = System::DynamicCast<Row>(table->get_LastRow()->Clone(true));
        // Remove all content from the cloned row's cells. This makes the row ready for new content to be inserted into.
        for (auto cell : System::IterateOver<Cell>(clonedRow->get_Cells()))
        {
            cell->RemoveAllChildren();
        }

        table->AppendChild(clonedRow);

        doc->Save(ArtifactsDir + u"WorkingWithTables.CloneLastRow.docx");
        //ExEnd:CloneLastRow
    }

    void FindingIndex()
    {
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        //ExStart:RetrieveTableIndex
        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        SharedPtr<NodeCollection> allTables = doc->GetChildNodes(NodeType::Table, true);
        int tableIndex = allTables->IndexOf(table);
        //ExEnd:RetrieveTableIndex
        std::cout << (String(u"\nTable index is ") + tableIndex) << std::endl;

        //ExStart:RetrieveRowIndex
        int rowIndex = table->IndexOf(table->get_LastRow());
        //ExEnd:RetrieveRowIndex
        std::cout << (String(u"\nRow index is ") + rowIndex) << std::endl;

        SharedPtr<Row> row = table->get_LastRow();
        //ExStart:RetrieveCellIndex
        int cellIndex = row->IndexOf(row->get_Cells()->idx_get(4));
        //ExEnd:RetrieveCellIndex
        std::cout << (String(u"\nCell index is ") + cellIndex) << std::endl;
    }

    void InsertTableDirectly()
    {
        //ExStart:InsertTableDirectly
        auto doc = MakeObject<Document>();

        // We start by creating the table object. Note that we must pass the document object
        // to the constructor of each node. This is because every node we create must belong
        // to some document.
        auto table = MakeObject<Table>(doc);
        doc->get_FirstSection()->get_Body()->AppendChild(table);

        // Here we could call EnsureMinimum to create the rows and cells for us. This method is used
        // to ensure that the specified node is valid. In this case, a valid table should have at least one Row and one cell.

        // Instead, we will handle creating the row and table ourselves.
        // This would be the best way to do this if we were creating a table inside an algorithm.
        auto row = MakeObject<Row>(doc);
        row->get_RowFormat()->set_AllowBreakAcrossPages(true);
        table->AppendChild(row);

        // We can now apply any auto fit settings.
        table->AutoFit(AutoFitBehavior::FixedColumnWidths);

        auto cell = MakeObject<Cell>(doc);
        cell->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightBlue());
        cell->get_CellFormat()->set_Width(80);
        cell->AppendChild(MakeObject<Paragraph>(doc));
        cell->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Row 1, Cell 1 Text"));

        row->AppendChild(cell);

        // We would then repeat the process for the other cells and rows in the table.
        // We can also speed things up by cloning existing cells and rows.
        row->AppendChild(cell->Clone(false));
        row->get_LastCell()->AppendChild(MakeObject<Paragraph>(doc));
        row->get_LastCell()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Row 1, Cell 2 Text"));

        doc->Save(ArtifactsDir + u"WorkingWithTables.InsertTableDirectly.docx");
        //ExEnd:InsertTableDirectly
    }

    void InsertTableFromHtml()
    {
        //ExStart:InsertTableFromHtml
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Note that AutoFitSettings does not apply to tables inserted from HTML.
        builder->InsertHtml(String(u"<table>") + u"<tr>" + u"<td>Row 1, Cell 1</td>" + u"<td>Row 1, Cell 2</td>" + u"</tr>" + u"<tr>" +
                            u"<td>Row 2, Cell 2</td>" + u"<td>Row 2, Cell 2</td>" + u"</tr>" + u"</table>");

        doc->Save(ArtifactsDir + u"WorkingWithTables.InsertTableFromHtml.docx");
        //ExEnd:InsertTableFromHtml
    }

    void CreateSimpleTable()
    {
        //ExStart:CreateSimpleTable
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Start building the table.
        builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 1 Content.");

        // Build the second cell.
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 2 Content.");

        // Call the following method to end the row and start a new row.
        builder->EndRow();

        // Build the first cell of the second row.
        builder->InsertCell();
        builder->Write(u"Row 2, Cell 1 Content");

        // Build the second cell.
        builder->InsertCell();
        builder->Write(u"Row 2, Cell 2 Content.");
        builder->EndRow();

        // Signal that we have finished building the table.
        builder->EndTable();

        doc->Save(ArtifactsDir + u"WorkingWithTables.CreateSimpleTable.docx");
        //ExEnd:CreateSimpleTable
    }

    void FormattedTable()
    {
        //ExStart:FormattedTable
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();

        // Table wide formatting must be applied after at least one row is present in the table.
        table->set_LeftIndent(20.0);

        // Set height and define the height rule for the header row.
        builder->get_RowFormat()->set_Height(40.0);
        builder->get_RowFormat()->set_HeightRule(HeightRule::AtLeast);

        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::FromArgb(198, 217, 241));
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
        builder->get_Font()->set_Size(16);
        builder->get_Font()->set_Name(u"Arial");
        builder->get_Font()->set_Bold(true);

        builder->get_CellFormat()->set_Width(100.0);
        builder->Write(u"Header Row,\n Cell 1");

        // We don't need to specify this cell's width because it's inherited from the previous cell.
        builder->InsertCell();
        builder->Write(u"Header Row,\n Cell 2");

        builder->InsertCell();
        builder->get_CellFormat()->set_Width(200.0);
        builder->Write(u"Header Row,\n Cell 3");
        builder->EndRow();

        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_White());
        builder->get_CellFormat()->set_Width(100.0);
        builder->get_CellFormat()->set_VerticalAlignment(CellVerticalAlignment::Center);

        // Reset height and define a different height rule for table body.
        builder->get_RowFormat()->set_Height(30.0);
        builder->get_RowFormat()->set_HeightRule(HeightRule::Auto);
        builder->InsertCell();

        // Reset font formatting.
        builder->get_Font()->set_Size(12);
        builder->get_Font()->set_Bold(false);

        builder->Write(u"Row 1, Cell 1 Content");
        builder->InsertCell();
        builder->Write(u"Row 1, Cell 2 Content");

        builder->InsertCell();
        builder->get_CellFormat()->set_Width(200.0);
        builder->Write(u"Row 1, Cell 3 Content");
        builder->EndRow();

        builder->InsertCell();
        builder->get_CellFormat()->set_Width(100.0);
        builder->Write(u"Row 2, Cell 1 Content");

        builder->InsertCell();
        builder->Write(u"Row 2, Cell 2 Content");

        builder->InsertCell();
        builder->get_CellFormat()->set_Width(200.0);
        builder->Write(u"Row 2, Cell 3 Content.");
        builder->EndRow();
        builder->EndTable();

        doc->Save(ArtifactsDir + u"WorkingWithTables.FormattedTable.docx");
        //ExEnd:FormattedTable
    }

    void NestedTable()
    {
        //ExStart:NestedTable
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Cell> cell = builder->InsertCell();
        builder->Writeln(u"Outer Table Cell 1");

        builder->InsertCell();
        builder->Writeln(u"Outer Table Cell 2");

        // This call is important to create a nested table within the first table.
        // Without this call, the cells inserted below will be appended to the outer table.
        builder->EndTable();

        // Move to the first cell of the outer table.
        builder->MoveTo(cell->get_FirstParagraph());

        // Build the inner table.
        builder->InsertCell();
        builder->Writeln(u"Inner Table Cell 1");
        builder->InsertCell();
        builder->Writeln(u"Inner Table Cell 2");
        builder->EndTable();

        doc->Save(ArtifactsDir + u"WorkingWithTables.NestedTable.docx");
        //ExEnd:NestedTable
    }

    void CombineRows()
    {
        //ExStart:CombineRows
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        // The rows from the second table will be appended to the end of the first table.
        auto firstTable = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        auto secondTable = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 1, true));

        // Append all rows from the current table to the next tables
        // with different cell count and widths can be joined into one table.
        while (secondTable->get_HasChildNodes())
        {
            firstTable->get_Rows()->Add(secondTable->get_FirstRow());
        }

        secondTable->Remove();

        doc->Save(ArtifactsDir + u"WorkingWithTables.CombineRows.docx");
        //ExEnd:CombineRows
    }

    void SplitTable()
    {
        //ExStart:SplitTable
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto firstTable = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        // We will split the table at the third row (inclusive).
        SharedPtr<Row> row = firstTable->get_Rows()->idx_get(2);

        // Create a new container for the split table.
        auto table = System::DynamicCast<Table>(firstTable->Clone(false));

        // Insert the container after the original.
        firstTable->get_ParentNode()->InsertAfter(table, firstTable);

        // Add a buffer paragraph to ensure the tables stay apart.
        firstTable->get_ParentNode()->InsertAfter(MakeObject<Paragraph>(doc), firstTable);

        SharedPtr<Row> currentRow;

        do
        {
            currentRow = firstTable->get_LastRow();
            table->PrependChild(currentRow);
        } while (currentRow != row);

        doc->Save(ArtifactsDir + u"WorkingWithTables.SplitTable.docx");
        //ExEnd:SplitTable
    }

    void RowFormatDisableBreakAcrossPages()
    {
        //ExStart:RowFormatDisableBreakAcrossPages
        auto doc = MakeObject<Document>(MyDir + u"Table spanning two pages.docx");

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        // Disable breaking across pages for all rows in the table.
        for (auto row : System::IterateOver<Row>(table->get_Rows()))
        {
            row->get_RowFormat()->set_AllowBreakAcrossPages(false);
        }

        doc->Save(ArtifactsDir + u"WorkingWithTables.RowFormatDisableBreakAcrossPages.docx");
        //ExEnd:RowFormatDisableBreakAcrossPages
    }

    void KeepTableTogether()
    {
        //ExStart:KeepTableTogether
        auto doc = MakeObject<Document>(MyDir + u"Table spanning two pages.docx");

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        // We need to enable KeepWithNext for every paragraph in the table to keep it from breaking across a page,
        // except for the last paragraphs in the last row of the table.
        for (auto cell : System::IterateOver<Cell>(table->GetChildNodes(NodeType::Cell, true)))
        {
            cell->EnsureMinimum();

            for (auto para : System::IterateOver<Paragraph>(cell->get_Paragraphs()))
            {
                if (!(cell->get_ParentRow()->get_IsLastRow() && para->get_IsEndOfCell()))
                {
                    para->get_ParagraphFormat()->set_KeepWithNext(true);
                }
            }
        }

        doc->Save(ArtifactsDir + u"WorkingWithTables.KeepTableTogether.docx");
        //ExEnd:KeepTableTogether
    }

    void CheckCellsMerged()
    {
        //ExStart:CheckCellsMerged
        auto doc = MakeObject<Document>(MyDir + u"Table with merged cells.docx");

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        for (auto row : System::IterateOver<Row>(table->get_Rows()))
        {
            for (auto cell : System::IterateOver<Cell>(row->get_Cells()))
            {
                std::cout << PrintCellMergeType(cell) << std::endl;
            }
        }
        //ExEnd:CheckCellsMerged
    }

    String PrintCellMergeType(SharedPtr<Cell> cell)
    {
        bool isHorizontallyMerged = cell->get_CellFormat()->get_HorizontalMerge() != CellMerge::None;
        bool isVerticallyMerged = cell->get_CellFormat()->get_VerticalMerge() != CellMerge::None;

        String cellLocation = String::Format(u"R{0}, C{1}", cell->get_ParentRow()->get_ParentTable()->IndexOf(cell->get_ParentRow()) + 1,
                                             cell->get_ParentRow()->IndexOf(cell) + 1);

        if (isHorizontallyMerged && isVerticallyMerged)
        {
            return String::Format(u"The cell at {0} is both horizontally and vertically merged", cellLocation);
        }

        if (isHorizontallyMerged)
        {
            return String::Format(u"The cell at {0} is horizontally merged.", cellLocation);
        }

        if (isVerticallyMerged)
        {
            return String::Format(u"The cell at {0} is vertically merged", cellLocation);
        }

        return String::Format(u"The cell at {0} is not merged", cellLocation);
    }

    void VerticalMerge()
    {
        //ExStart:VerticalMerge
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

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

        doc->Save(ArtifactsDir + u"WorkingWithTables.VerticalMerge.docx");
        //ExEnd:VerticalMerge
    }

    void HorizontalMerge()
    {
        //ExStart:HorizontalMerge
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

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

        doc->Save(ArtifactsDir + u"WorkingWithTables.HorizontalMerge.docx");
        //ExEnd:HorizontalMerge
    }

    void MergeCellRange()
    {
        //ExStart:MergeCellRange
        auto doc = MakeObject<Document>(MyDir + u"Table with merged cells.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        // We want to merge the range of cells found inbetween these two cells.
        SharedPtr<Cell> cellStartRange = table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0);
        SharedPtr<Cell> cellEndRange = table->get_Rows()->idx_get(1)->get_Cells()->idx_get(1);

        // Merge all the cells between the two specified cells into one.
        MergeCells(cellStartRange, cellEndRange);

        doc->Save(ArtifactsDir + u"WorkingWithTables.MergeCellRange.docx");
        //ExEnd:MergeCellRange
    }

    void PrintHorizontalAndVerticalMerged()
    {
        //ExStart:PrintHorizontalAndVerticalMerged
        auto doc = MakeObject<Document>(MyDir + u"Table with merged cells.docx");

        auto visitor = MakeObject<WorkingWithTables::SpanVisitor>(doc);
        doc->Accept(visitor);
        //ExEnd:PrintHorizontalAndVerticalMerged
    }

    void ConvertToHorizontallyMergedCells()
    {
        //ExStart:ConvertToHorizontallyMergedCells
        auto doc = MakeObject<Document>(MyDir + u"Table with merged cells.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        // Now merged cells have appropriate merge flags.
        table->ConvertToHorizontallyMergedCells();
        //ExEnd:ConvertToHorizontallyMergedCells
    }

    void MergeCells(SharedPtr<Cell> startCell, SharedPtr<Cell> endCell)
    {
        SharedPtr<Table> parentTable = startCell->get_ParentRow()->get_ParentTable();

        // Find the row and cell indices for the start and end cell.
        System::Drawing::Point startCellPos(startCell->get_ParentRow()->IndexOf(startCell), parentTable->IndexOf(startCell->get_ParentRow()));
        System::Drawing::Point endCellPos(endCell->get_ParentRow()->IndexOf(endCell), parentTable->IndexOf(endCell->get_ParentRow()));

        // Create a range of cells to be merged based on these indices.
        // Inverse each index if the end cell is before the start cell.
        System::Drawing::Rectangle mergeRange(
            System::Math::Min(startCellPos.get_X(), endCellPos.get_X()), System::Math::Min(startCellPos.get_Y(), endCellPos.get_Y()),
            System::Math::Abs(endCellPos.get_X() - startCellPos.get_X()) + 1, System::Math::Abs(endCellPos.get_Y() - startCellPos.get_Y()) + 1);

        for (auto row : System::IterateOver<Row>(parentTable->get_Rows()))
        {
            for (auto cell : System::IterateOver<Cell>(row->get_Cells()))
            {
                System::Drawing::Point currentPos(row->IndexOf(cell), parentTable->IndexOf(row));

                // Check if the current cell is inside our merge range, then merge it.
                if (mergeRange.Contains(currentPos))
                {
                    cell->get_CellFormat()->set_HorizontalMerge(currentPos.get_X() == mergeRange.get_X() ? CellMerge::First : CellMerge::Previous);

                    cell->get_CellFormat()->set_VerticalMerge(currentPos.get_Y() == mergeRange.get_Y() ? CellMerge::First : CellMerge::Previous);
                }
            }
        }
    }

    void RepeatRowsOnSubsequentPages()
    {
        //ExStart:RepeatRowsOnSubsequentPages
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartTable();
        builder->get_RowFormat()->set_HeadingFormat(true);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
        builder->get_CellFormat()->set_Width(100);
        builder->InsertCell();
        builder->Writeln(u"Heading row 1");
        builder->EndRow();
        builder->InsertCell();
        builder->Writeln(u"Heading row 2");
        builder->EndRow();

        builder->get_CellFormat()->set_Width(50);
        builder->get_ParagraphFormat()->ClearFormatting();

        for (int i = 0; i < 50; i++)
        {
            builder->InsertCell();
            builder->get_RowFormat()->set_HeadingFormat(false);
            builder->Write(u"Column 1 Text");
            builder->InsertCell();
            builder->Write(u"Column 2 Text");
            builder->EndRow();
        }

        doc->Save(ArtifactsDir + u"WorkingWithTables.RepeatRowsOnSubsequentPages.docx");
        //ExEnd:RepeatRowsOnSubsequentPages
    }

    void AutoFitToPageWidth()
    {
        //ExStart:AutoFitToPageWidth
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a table with a width that takes up half the page width.
        SharedPtr<Table> table = builder->StartTable();

        builder->InsertCell();
        table->set_PreferredWidth(PreferredWidth::FromPercent(50));
        builder->Writeln(u"Cell #1");

        builder->InsertCell();
        builder->Writeln(u"Cell #2");

        builder->InsertCell();
        builder->Writeln(u"Cell #3");

        doc->Save(ArtifactsDir + u"WorkingWithTables.AutoFitToPageWidth.docx");
        //ExEnd:AutoFitToPageWidth
    }

    void PreferredWidthSettings()
    {
        //ExStart:PreferredWidthSettings
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a table row made up of three cells which have different preferred widths.
        builder->StartTable();

        // Insert an absolute sized cell.
        builder->InsertCell();
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPoints(40));
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightYellow());
        builder->Writeln(u"Cell at 40 points width");

        // Insert a relative (percent) sized cell.
        builder->InsertCell();
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPercent(20));
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightBlue());
        builder->Writeln(u"Cell at 20% width");

        // Insert a auto sized cell.
        builder->InsertCell();
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::Auto());
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightGreen());
        builder->Writeln(u"Cell automatically sized. The size of this cell is calculated from the table preferred width.");
        builder->Writeln(u"In this case the cell will fill up the rest of the available space.");

        doc->Save(ArtifactsDir + u"WorkingWithTables.PreferredWidthSettings.docx");
        //ExEnd:PreferredWidthSettings
    }

    void RetrievePreferredWidthType()
    {
        //ExStart:RetrievePreferredWidthType
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        //ExStart:AllowAutoFit
        table->set_AllowAutoFit(true);
        //ExEnd:AllowAutoFit

        SharedPtr<Cell> firstCell = table->get_FirstRow()->get_FirstCell();
        PreferredWidthType type = firstCell->get_CellFormat()->get_PreferredWidth()->get_Type();
        double value = firstCell->get_CellFormat()->get_PreferredWidth()->get_Value();
        //ExEnd:RetrievePreferredWidthType
    }

    void GetTablePosition()
    {
        //ExStart:GetTablePosition
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        if (table->get_TextWrapping() == TextWrapping::Around)
        {
            std::cout << System::EnumGetName(table->get_RelativeHorizontalAlignment()) << std::endl;
            std::cout << System::EnumGetName(table->get_RelativeVerticalAlignment()) << std::endl;
        }
        else
        {
            std::cout << System::EnumGetName(table->get_Alignment()) << std::endl;
        }
        //ExEnd:GetTablePosition
    }

    void GetFloatingTablePosition()
    {
        //ExStart:GetFloatingTablePosition
        auto doc = MakeObject<Document>(MyDir + u"Table wrapped by text.docx");

        for (auto table : System::IterateOver<Table>(doc->get_FirstSection()->get_Body()->get_Tables()))
        {
            // If the table is floating type, then print its positioning properties.
            if (table->get_TextWrapping() == TextWrapping::Around)
            {
                std::cout << System::EnumGetName(table->get_HorizontalAnchor()) << std::endl;
                std::cout << System::EnumGetName(table->get_VerticalAnchor()) << std::endl;
                std::cout << table->get_AbsoluteHorizontalDistance() << std::endl;
                std::cout << table->get_AbsoluteVerticalDistance() << std::endl;
                std::cout << System::Convert::ToString(table->get_AllowOverlap()) << std::endl;
                std::cout << table->get_AbsoluteHorizontalDistance() << std::endl;
                std::cout << System::EnumGetName(table->get_RelativeVerticalAlignment()) << std::endl;
                std::cout << ".............................." << std::endl;
            }
        }
        //ExEnd:GetFloatingTablePosition
    }

    void FloatingTablePosition()
    {
        //ExStart:FloatingTablePosition
        auto doc = MakeObject<Document>(MyDir + u"Table wrapped by text.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        table->set_AbsoluteHorizontalDistance(10);
        table->set_RelativeVerticalAlignment(VerticalAlignment::Center);

        doc->Save(ArtifactsDir + u"WorkingWithTables.FloatingTablePosition.docx");
        //ExEnd:FloatingTablePosition
    }

    void SetRelativeHorizontalOrVerticalPosition()
    {
        //ExStart:SetRelativeHorizontalOrVerticalPosition
        auto doc = MakeObject<Document>(MyDir + u"Table wrapped by text.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        table->set_HorizontalAnchor(RelativeHorizontalPosition::Column);
        table->set_VerticalAnchor(RelativeVerticalPosition::Page);

        doc->Save(ArtifactsDir + u"WorkingWithTables.SetFloatingTablePosition.docx");
        //ExEnd:SetRelativeHorizontalOrVerticalPosition
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Tables
