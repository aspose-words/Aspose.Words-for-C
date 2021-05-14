#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Border.h>
#include <Aspose.Words.Cpp/BorderCollection.h>
#include <Aspose.Words.Cpp/BorderType.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/ConditionalStyle.h>
#include <Aspose.Words.Cpp/ConditionalStyleCollection.h>
#include <Aspose.Words.Cpp/ConditionalStyleType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/HeightRule.h>
#include <Aspose.Words.Cpp/LineStyle.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Replacing/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Shading.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleType.h>
#include <Aspose.Words.Cpp/TableStyle.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Tables/CellMerge.h>
#include <Aspose.Words.Cpp/Tables/CellVerticalAlignment.h>
#include <Aspose.Words.Cpp/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Tables/PreferredWidthType.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/Tables/TableAlignment.h>
#include <Aspose.Words.Cpp/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Tables/TableStyleOptions.h>
#include <Aspose.Words.Cpp/Tables/TextWrapping.h>
#include <Aspose.Words.Cpp/TextOrientation.h>
#include <Aspose.Words.Cpp/TextureIndex.h>
#include <drawing/color.h>
#include <drawing/point.h>
#include <drawing/rectangle.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ienumerator.h>
#include <system/details/dispose_guard.h>
#include <system/enum.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/linq/enumerable.h>
#include <system/math.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/regularexpressions/regex.h>
#include <system/type_info.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Tables;

namespace ApiExamples {

class ExTable : public ApiExampleBase
{
public:
    void CreateTable()
    {
        //ExStart
        //ExFor:Table
        //ExFor:Row
        //ExFor:Cell
        //ExFor:Table.#ctor(DocumentBase)
        //ExSummary:Shows how to create a table.
        auto doc = MakeObject<Document>();
        auto table = MakeObject<Table>(doc);
        doc->get_FirstSection()->get_Body()->AppendChild(table);

        // Tables contain rows, which contain cells, which may have paragraphs
        // with typical elements such as runs, shapes, and even other tables.
        // Calling the "EnsureMinimum" method on a table will ensure that
        // the table has at least one row, cell, and paragraph.
        auto firstRow = MakeObject<Row>(doc);
        table->AppendChild(firstRow);

        auto firstCell = MakeObject<Cell>(doc);
        firstRow->AppendChild(firstCell);

        auto paragraph = MakeObject<Paragraph>(doc);
        firstCell->AppendChild(paragraph);

        // Add text to the first call in the first row of the table.
        auto run = MakeObject<Run>(doc, u"Hello world!");
        paragraph->AppendChild(run);

        doc->Save(ArtifactsDir + u"Table.CreateTable.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.CreateTable.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(1, table->get_Rows()->get_Count());
        ASSERT_EQ(1, table->get_FirstRow()->get_Cells()->get_Count());
        ASSERT_EQ(u"Hello world!\a\a", table->GetText().Trim());
    }

    void Padding()
    {
        //ExStart
        //ExFor:Table.LeftPadding
        //ExFor:Table.RightPadding
        //ExFor:Table.TopPadding
        //ExFor:Table.BottomPadding
        //ExSummary:Shows how to configure content padding in a table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Row 1, cell 1.");
        builder->InsertCell();
        builder->Write(u"Row 1, cell 2.");
        builder->EndTable();

        // For every cell in the table, set the distance between its contents and each of its borders.
        // This table will maintain the minimum padding distance by wrapping text.
        table->set_LeftPadding(30);
        table->set_RightPadding(60);
        table->set_TopPadding(10);
        table->set_BottomPadding(90);
        table->set_PreferredWidth(PreferredWidth::FromPoints(250));

        doc->Save(ArtifactsDir + u"DocumentBuilder.SetRowFormatting.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBuilder.SetRowFormatting.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASPOSE_ASSERT_EQ(30.0, table->get_LeftPadding());
        ASPOSE_ASSERT_EQ(60.0, table->get_RightPadding());
        ASPOSE_ASSERT_EQ(10.0, table->get_TopPadding());
        ASPOSE_ASSERT_EQ(90.0, table->get_BottomPadding());
    }

    void RowCellFormat()
    {
        //ExStart
        //ExFor:Row.RowFormat
        //ExFor:RowFormat
        //ExFor:Cell.CellFormat
        //ExFor:CellFormat
        //ExFor:CellFormat.Shading
        //ExSummary:Shows how to modify the format of rows and cells in a table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"City");
        builder->InsertCell();
        builder->Write(u"Country");
        builder->EndRow();
        builder->InsertCell();
        builder->Write(u"London");
        builder->InsertCell();
        builder->Write(u"U.K.");
        builder->EndTable();

        // Use the first row's "RowFormat" property to modify the formatting
        // of the contents of all cells in this row.
        SharedPtr<RowFormat> rowFormat = table->get_FirstRow()->get_RowFormat();
        rowFormat->set_Height(25);
        rowFormat->get_Borders()->idx_get(BorderType::Bottom)->set_Color(System::Drawing::Color::get_Red());

        // Use the "CellFormat" property of the first cell in the last row to modify the formatting of that cell's contents.
        SharedPtr<CellFormat> cellFormat = table->get_LastRow()->get_FirstCell()->get_CellFormat();
        cellFormat->set_Width(100);
        cellFormat->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Orange());

        doc->Save(ArtifactsDir + u"Table.RowCellFormat.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.RowCellFormat.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(u"City\aCountry\a\aLondon\aU.K.\a\a", table->GetText().Trim());

        rowFormat = table->get_FirstRow()->get_RowFormat();

        ASPOSE_ASSERT_EQ(25.0, rowFormat->get_Height());
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), rowFormat->get_Borders()->idx_get(BorderType::Bottom)->get_Color().ToArgb());

        cellFormat = table->get_LastRow()->get_FirstCell()->get_CellFormat();

        ASPOSE_ASSERT_EQ(110.8, cellFormat->get_Width());
        ASSERT_EQ(System::Drawing::Color::get_Orange().ToArgb(), cellFormat->get_Shading()->get_BackgroundPatternColor().ToArgb());
    }

    void DisplayContentOfTables()
    {
        //ExStart
        //ExFor:Cell
        //ExFor:CellCollection
        //ExFor:CellCollection.Item(System.Int32)
        //ExFor:CellCollection.ToArray
        //ExFor:Row
        //ExFor:Row.Cells
        //ExFor:RowCollection
        //ExFor:RowCollection.Item(System.Int32)
        //ExFor:RowCollection.ToArray
        //ExFor:Table
        //ExFor:Table.Rows
        //ExFor:TableCollection.Item(System.Int32)
        //ExFor:TableCollection.ToArray
        //ExSummary:Shows how to iterate through all tables in the document and print the contents of each cell.
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
        SharedPtr<TableCollection> tables = doc->get_FirstSection()->get_Body()->get_Tables();

        ASSERT_EQ(2, tables->ToArray()->get_Length());

        for (int i = 0; i < tables->get_Count(); i++)
        {
            std::cout << "Start of Table " << i << std::endl;

            SharedPtr<RowCollection> rows = tables->idx_get(i)->get_Rows();

            // We can use the "ToArray" method on a row collection to clone it into an array.
            ASPOSE_ASSERT_EQ(rows, rows->ToArray());
            ASPOSE_ASSERT_NS(rows, rows->ToArray());

            for (int j = 0; j < rows->get_Count(); j++)
            {
                std::cout << "\tStart of Row " << j << std::endl;

                SharedPtr<CellCollection> cells = rows->idx_get(j)->get_Cells();

                // We can use the "ToArray" method on a cell collection to clone it into an array.
                ASPOSE_ASSERT_EQ(cells, cells->ToArray());
                ASPOSE_ASSERT_NS(cells, cells->ToArray());

                for (int k = 0; k < cells->get_Count(); k++)
                {
                    String cellText = cells->idx_get(k)->ToString(SaveFormat::Text).Trim();
                    std::cout << "\t\tContents of Cell:" << k << " = \"" << cellText << "\"" << std::endl;
                }

                std::cout << "\tEnd of Row " << j << std::endl;
            }

            std::cout << "End of Table " << i << "\n" << std::endl;
        }
        //ExEnd
    }

    //ExStart
    //ExFor:Node.GetAncestor(NodeType)
    //ExFor:Node.GetAncestor(System.Type)
    //ExFor:Table.NodeType
    //ExFor:Cell.Tables
    //ExFor:TableCollection
    //ExFor:NodeCollection.Count
    //ExSummary:Shows how to find out if a tables are nested.
    void CalculateDepthOfNestedTables()
    {
        auto doc = MakeObject<Document>(MyDir + u"Nested tables.docx");
        SharedPtr<NodeCollection> tables = doc->GetChildNodes(NodeType::Table, true);
        ASSERT_EQ(5, tables->get_Count());
        //ExSkip

        for (int i = 0; i < tables->get_Count(); i++)
        {
            auto table = System::DynamicCast<Table>(tables->idx_get(i));

            // Find out if any cells in the table have other tables as children.
            int count = GetChildTableCount(table);
            std::cout << "Table #" << i << " has " << count << " tables directly within its cells" << std::endl;

            // Find out if the table is nested inside another table, and, if so, at what depth.
            int tableDepth = GetNestedDepthOfTable(table);

            if (tableDepth > 0)
            {
                std::cout << "Table #" << i << " is nested inside another table at depth of " << tableDepth << std::endl;
            }
            else
            {
                std::cout << "Table #" << i << " is a non nested table (is not a child of another table)" << std::endl;
            }
        }
    }

    /// <summary>
    /// Calculates what level a table is nested inside other tables.
    /// </summary>
    /// <returns>
    /// An integer indicating the nesting depth of the table (number of parent table nodes).
    /// </returns>
    static int GetNestedDepthOfTable(SharedPtr<Table> table)
    {
        int depth = 0;
        SharedPtr<Node> parent = table->GetAncestor(table->get_NodeType());

        while (parent != nullptr)
        {
            depth++;
            parent = parent->GetAncestorOf<SharedPtr<Table>>();
        }

        return depth;
    }

    /// <summary>
    /// Determines if a table contains any immediate child table within its cells.
    /// Do not recursively traverse through those tables to check for further tables.
    /// </summary>
    /// <returns>
    /// Returns true if at least one child cell contains a table.
    /// Returns false if no cells in the table contain a table.
    /// </returns>
    static int GetChildTableCount(SharedPtr<Table> table)
    {
        int childTableCount = 0;

        for (auto row : System::IterateOver(table->get_Rows()->LINQ_OfType<SharedPtr<Row>>()))
        {
            for (auto Cell : System::IterateOver(row->get_Cells()->LINQ_OfType<SharedPtr<Cell>>()))
            {
                SharedPtr<TableCollection> childTables = Cell->get_Tables();

                if (childTables->get_Count() > 0)
                {
                    childTableCount++;
                }
            }
        }

        return childTableCount;
    }
    //ExEnd

    void EnsureTableMinimum()
    {
        //ExStart
        //ExFor:Table.EnsureMinimum
        //ExSummary:Shows how to ensure that a table node contains the nodes we need to add content.
        auto doc = MakeObject<Document>();
        auto table = MakeObject<Table>(doc);
        doc->get_FirstSection()->get_Body()->AppendChild(table);

        // Tables contain rows, which contain cells, which may contain paragraphs
        // with typical elements such as runs, shapes, and even other tables.
        // Our new table has none of these nodes, and we cannot add contents to it until it does.
        ASSERT_EQ(0, table->GetChildNodes(NodeType::Any, true)->get_Count());

        // Calling the "EnsureMinimum" method on a table will ensure that
        // the table has at least one row and one cell with an empty paragraph.
        table->EnsureMinimum();
        table->get_FirstRow()->get_FirstCell()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Hello world!"));
        //ExEnd

        ASSERT_EQ(4, table->GetChildNodes(NodeType::Any, true)->get_Count());
    }

    void EnsureRowMinimum()
    {
        //ExStart
        //ExFor:Row.EnsureMinimum
        //ExSummary:Shows how to ensure a row node contains the nodes we need to begin adding content to it.
        auto doc = MakeObject<Document>();
        auto table = MakeObject<Table>(doc);
        doc->get_FirstSection()->get_Body()->AppendChild(table);
        auto row = MakeObject<Row>(doc);
        table->AppendChild(row);

        // Rows contain cells, containing paragraphs with typical elements such as runs, shapes, and even other tables.
        // Our new row has none of these nodes, and we cannot add contents to it until it does.
        ASSERT_EQ(0, row->GetChildNodes(NodeType::Any, true)->get_Count());

        // Calling the "EnsureMinimum" method on a table will ensure that
        // the table has at least one cell with an empty paragraph.
        row->EnsureMinimum();
        row->get_FirstCell()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Hello world!"));
        //ExEnd

        ASSERT_EQ(3, row->GetChildNodes(NodeType::Any, true)->get_Count());
    }

    void EnsureCellMinimum()
    {
        //ExStart
        //ExFor:Cell.EnsureMinimum
        //ExSummary:Shows how to ensure a cell node contains the nodes we need to begin adding content to it.
        auto doc = MakeObject<Document>();
        auto table = MakeObject<Table>(doc);
        doc->get_FirstSection()->get_Body()->AppendChild(table);
        auto row = MakeObject<Row>(doc);
        table->AppendChild(row);
        auto cell = MakeObject<Cell>(doc);
        row->AppendChild(cell);

        // Cells may contain paragraphs with typical elements such as runs, shapes, and even other tables.
        // Our new cell does not have any paragraphs, and we cannot add contents such as run and shape nodes to it until it does.
        ASSERT_EQ(0, cell->GetChildNodes(NodeType::Any, true)->get_Count());

        // Calling the "EnsureMinimum" method on a cell will ensure that
        // the cell has at least one empty paragraph, which we can then add contents to.
        cell->EnsureMinimum();
        cell->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Hello world!"));
        //ExEnd

        ASSERT_EQ(2, cell->GetChildNodes(NodeType::Any, true)->get_Count());
    }

    void SetOutlineBorders()
    {
        //ExStart
        //ExFor:Table.Alignment
        //ExFor:TableAlignment
        //ExFor:Table.ClearBorders
        //ExFor:Table.ClearShading
        //ExFor:Table.SetBorder
        //ExFor:TextureIndex
        //ExFor:Table.SetShading
        //ExSummary:Shows how to apply an outline border to a table.
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        // Align the table to the center of the page.
        table->set_Alignment(TableAlignment::Center);

        // Clear any existing borders and shading from the table.
        table->ClearBorders();
        table->ClearShading();

        // Add green borders to the outline of the table.
        table->SetBorder(BorderType::Left, LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
        table->SetBorder(BorderType::Right, LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
        table->SetBorder(BorderType::Top, LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
        table->SetBorder(BorderType::Bottom, LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);

        // Fill the cells with a light green solid color.
        table->SetShading(TextureIndex::TextureSolid, System::Drawing::Color::get_LightGreen(), System::Drawing::Color::Empty);

        doc->Save(ArtifactsDir + u"Table.SetOutlineBorders.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.SetOutlineBorders.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(TableAlignment::Center, table->get_Alignment());

        SharedPtr<BorderCollection> borders = table->get_FirstRow()->get_RowFormat()->get_Borders();

        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), borders->get_Top()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), borders->get_Left()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), borders->get_Right()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), borders->get_Bottom()->get_Color().ToArgb());
        ASSERT_NE(System::Drawing::Color::get_Green().ToArgb(), borders->get_Horizontal()->get_Color().ToArgb());
        ASSERT_NE(System::Drawing::Color::get_Green().ToArgb(), borders->get_Vertical()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_LightGreen().ToArgb(),
                  table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_ForegroundPatternColor().ToArgb());
    }

    void SetBorders()
    {
        //ExStart
        //ExFor:Table.SetBorders
        //ExSummary:Shows how to format of all of a table's borders at once.
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        // Clear all existing borders from the table.
        table->ClearBorders();

        // Set a single green line to serve as every outer and inner border of this table.
        table->SetBorders(LineStyle::Single, 1.5, System::Drawing::Color::get_Green());

        doc->Save(ArtifactsDir + u"Table.SetBorders.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.SetBorders.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Top()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Left()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Right()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Bottom()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Horizontal()->get_Color().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Vertical()->get_Color().ToArgb());
    }

    void RowFormat_()
    {
        //ExStart
        //ExFor:RowFormat
        //ExFor:Row.RowFormat
        //ExSummary:Shows how to modify formatting of a table row.
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        // Use the first row's "RowFormat" property to set formatting that modifies that entire row's appearance.
        SharedPtr<Row> firstRow = table->get_FirstRow();
        firstRow->get_RowFormat()->get_Borders()->set_LineStyle(LineStyle::None);
        firstRow->get_RowFormat()->set_HeightRule(HeightRule::Auto);
        firstRow->get_RowFormat()->set_AllowBreakAcrossPages(true);

        doc->Save(ArtifactsDir + u"Table.RowFormat.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.RowFormat.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(LineStyle::None, table->get_FirstRow()->get_RowFormat()->get_Borders()->get_LineStyle());
        ASSERT_EQ(HeightRule::Auto, table->get_FirstRow()->get_RowFormat()->get_HeightRule());
        ASSERT_TRUE(table->get_FirstRow()->get_RowFormat()->get_AllowBreakAcrossPages());
    }

    void CellFormat_()
    {
        //ExStart
        //ExFor:CellFormat
        //ExFor:Cell.CellFormat
        //ExSummary:Shows how to modify formatting of a table cell.
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        SharedPtr<Cell> firstCell = table->get_FirstRow()->get_FirstCell();

        // Use a cell's "CellFormat" property to set formatting that modifies the appearance of that cell.
        firstCell->get_CellFormat()->set_Width(30);
        firstCell->get_CellFormat()->set_Orientation(TextOrientation::Downward);
        firstCell->get_CellFormat()->get_Shading()->set_ForegroundPatternColor(System::Drawing::Color::get_LightGreen());

        doc->Save(ArtifactsDir + u"Table.CellFormat.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.CellFormat.docx");

        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        ASPOSE_ASSERT_EQ(30, table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Width());
        ASSERT_EQ(TextOrientation::Downward, table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Orientation());
        ASSERT_EQ(System::Drawing::Color::get_LightGreen().ToArgb(),
                  table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_ForegroundPatternColor().ToArgb());
    }

    void GetDistance()
    {
        //ExStart
        //ExFor:Table.DistanceBottom
        //ExFor:Table.DistanceLeft
        //ExFor:Table.DistanceRight
        //ExFor:Table.DistanceTop
        //ExSummary:Shows the minimum distance operations between table boundaries and text.
        auto doc = MakeObject<Document>(MyDir + u"Table wrapped by text.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASPOSE_ASSERT_EQ(25.9, table->get_DistanceTop());
        ASPOSE_ASSERT_EQ(25.9, table->get_DistanceBottom());
        ASPOSE_ASSERT_EQ(17.3, table->get_DistanceLeft());
        ASPOSE_ASSERT_EQ(17.3, table->get_DistanceRight());
        //ExEnd
    }

    void Borders()
    {
        //ExStart
        //ExFor:Table.ClearBorders
        //ExSummary:Shows how to remove all borders from a table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Hello world!");
        builder->EndTable();

        // Modify the color and thickness of the top border.
        SharedPtr<Border> topBorder = table->get_FirstRow()->get_RowFormat()->get_Borders()->idx_get(BorderType::Top);
        table->SetBorder(BorderType::Top, LineStyle::Double, 1.5, System::Drawing::Color::get_Red(), true);

        ASPOSE_ASSERT_EQ(1.5, topBorder->get_LineWidth());
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), topBorder->get_Color().ToArgb());
        ASSERT_EQ(LineStyle::Double, topBorder->get_LineStyle());

        // Clear the borders of all cells in the table, and then save the document.
        table->ClearBorders();
        ASSERT_NE(System::Drawing::Color::Empty.ToArgb(), topBorder->get_Color().ToArgb());
        //ExSkip
        doc->Save(ArtifactsDir + u"Table.ClearBorders.docx");

        // Verify the values of the table's properties after re-opening the document.
        doc = MakeObject<Document>(ArtifactsDir + u"Table.ClearBorders.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        topBorder = table->get_FirstRow()->get_RowFormat()->get_Borders()->idx_get(BorderType::Top);

        ASPOSE_ASSERT_EQ(0.0, topBorder->get_LineWidth());
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), topBorder->get_Color().ToArgb());
        ASSERT_EQ(LineStyle::None, topBorder->get_LineStyle());
        //ExEnd
    }

    void ReplaceCellText()
    {
        //ExStart
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExSummary:Shows how to replace all instances of String of text in a table and cell.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Carrots");
        builder->InsertCell();
        builder->Write(u"50");
        builder->EndRow();
        builder->InsertCell();
        builder->Write(u"Potatoes");
        builder->InsertCell();
        builder->Write(u"50");
        builder->EndTable();

        auto options = MakeObject<FindReplaceOptions>();
        options->set_MatchCase(true);
        options->set_FindWholeWordsOnly(true);

        // Perform a find-and-replace operation on an entire table.
        table->get_Range()->Replace(u"Carrots", u"Eggs", options);

        // Perform a find-and-replace operation on the last cell of the last row of the table.
        table->get_LastRow()->get_LastCell()->get_Range()->Replace(u"50", u"20", options);

        ASSERT_EQ(String(u"Eggs\a50\a\a") + u"Potatoes\a20\a\a", table->GetText().Trim());
        //ExEnd
    }

    void RemoveParagraphTextAndMark(bool isSmartParagraphBreakReplacement)
    {
        //ExStart
        //ExFor:FindReplaceOptions.SmartParagraphBreakReplacement
        //ExSummary:Shows how to remove paragraph from a table cell with a nested table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create table with paragraph and inner table in first cell.
        builder->StartTable();
        builder->InsertCell();
        builder->Write(u"TEXT1");
        builder->StartTable();
        builder->InsertCell();
        builder->EndTable();
        builder->EndTable();
        builder->Writeln();

        auto options = MakeObject<FindReplaceOptions>();
        // When the following option is set to 'true', Aspose.Words will remove paragraph's text
        // completely with its paragraph mark. Otherwise, Aspose.Words will mimic Word and remove
        // only paragraph's text and leaves the paragraph mark intact (when a table follows the text).
        options->set_SmartParagraphBreakReplacement(isSmartParagraphBreakReplacement);
        doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"TEXT1&p"), u"", options);

        doc->Save(ArtifactsDir + u"Table.RemoveParagraphTextAndMark.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.RemoveParagraphTextAndMark.docx");

        ASSERT_EQ(
            isSmartParagraphBreakReplacement ? 1 : 2,
            doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->get_Paragraphs()->get_Count());
    }

    void PrintTableRange()
    {
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        // The range text will include control characters such as "\a" for a cell.
        // You can call ToString on the desired node to retrieve the plain text content.

        // Print the plain text range of the table to the screen.
        std::cout << "Contents of the table: " << std::endl;
        std::cout << table->get_Range()->get_Text() << std::endl;

        // Print the contents of the second row to the screen.
        std::cout << "\nContents of the row: " << std::endl;
        std::cout << table->get_Rows()->idx_get(1)->get_Range()->get_Text() << std::endl;

        // Print the contents of the last cell in the table to the screen.
        std::cout << "\nContents of the cell: " << std::endl;
        std::cout << table->get_LastRow()->get_LastCell()->get_Range()->get_Text() << std::endl;

        ASSERT_EQ(u"\aColumn 1\aColumn 2\aColumn 3\aColumn 4\a\a", table->get_Rows()->idx_get(1)->get_Range()->get_Text());
        ASSERT_EQ(u"Cell 12 contents\a", table->get_LastRow()->get_LastCell()->get_Range()->get_Text());
    }

    void CloneTable()
    {
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        auto tableClone = System::DynamicCast<Table>(table->Clone(true));

        // Insert the cloned table into the document after the original.
        table->get_ParentNode()->InsertAfter(tableClone, table);

        // Insert an empty paragraph between the two tables.
        table->get_ParentNode()->InsertAfter(MakeObject<Paragraph>(doc), table);

        doc->Save(ArtifactsDir + u"Table.CloneTable.doc");

        ASSERT_EQ(3, doc->GetChildNodes(NodeType::Table, true)->get_Count());
        ASSERT_EQ(table->get_Range()->get_Text(), tableClone->get_Range()->get_Text());

        for (auto cell : System::IterateOver(tableClone->GetChildNodes(NodeType::Cell, true)->LINQ_OfType<SharedPtr<Cell>>()))
        {
            cell->RemoveAllChildren();
        }

        ASSERT_EQ(String::Empty, tableClone->ToString(SaveFormat::Text).Trim());
    }

    void AllowBreakAcrossPages(bool allowBreakAcrossPages)
    {
        //ExStart
        //ExFor:RowFormat.AllowBreakAcrossPages
        //ExSummary:Shows how to disable rows breaking across pages for every row in a table.
        auto doc = MakeObject<Document>(MyDir + u"Table spanning two pages.docx");
        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        // Set the "AllowBreakAcrossPages" property to "false" to keep the row
        // in one piece if a table spans two pages, which break up along that row.
        // If the row is too big to fit in one page, Microsoft Word will push it down to the next page.
        // Set the "AllowBreakAcrossPages" property to "true" to allow the row to break up across two pages.
        for (auto row : System::IterateOver(table->LINQ_OfType<SharedPtr<Row>>()))
        {
            row->get_RowFormat()->set_AllowBreakAcrossPages(allowBreakAcrossPages);
        }

        doc->Save(ArtifactsDir + u"Table.AllowBreakAcrossPages.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.AllowBreakAcrossPages.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(3, table->get_Rows()->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(static_cast<std::function<bool(SharedPtr<Node> r)>>(
                         [&allowBreakAcrossPages](SharedPtr<Node> r) -> bool
                         { return (System::DynamicCast<Row>(r))->get_RowFormat()->get_AllowBreakAcrossPages() == allowBreakAcrossPages; }))));
    }

    void AllowAutoFitOnTable(bool allowAutoFit)
    {
        //ExStart
        //ExFor:Table.AllowAutoFit
        //ExSummary:Shows how to enable/disable automatic table cell resizing.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPoints(100));
        builder->Write(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") +
                       u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        builder->InsertCell();
        builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::Auto());
        builder->Write(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") +
                       u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        builder->EndRow();
        builder->EndTable();

        // Set the "AllowAutoFit" property to "false" to get the table to maintain the dimensions
        // of all its rows and cells, and truncate contents if they get too large to fit.
        // Set the "AllowAutoFit" property to "true" to allow the table to change its cells' width and height
        // to accommodate their contents.
        table->set_AllowAutoFit(allowAutoFit);

        doc->Save(ArtifactsDir + u"Table.AllowAutoFitOnTable.html");
        //ExEnd

        if (allowAutoFit)
        {
            TestUtil::FileContainsString(u"<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; "
                                         u"padding-left:5.03pt; vertical-align:top; -aw-border-right:0.5pt single\">",
                                         ArtifactsDir + u"Table.AllowAutoFitOnTable.html");
            TestUtil::FileContainsString(u"<td style=\"border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; "
                                         u"vertical-align:top; -aw-border-left:0.5pt single\">",
                                         ArtifactsDir + u"Table.AllowAutoFitOnTable.html");
        }
        else
        {
            TestUtil::FileContainsString(u"<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; "
                                         u"padding-left:5.03pt; vertical-align:top; -aw-border-right:0.5pt single\">",
                                         ArtifactsDir + u"Table.AllowAutoFitOnTable.html");
            TestUtil::FileContainsString(u"<td style=\"width:7.2pt; border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; "
                                         u"padding-left:5.03pt; vertical-align:top; -aw-border-left:0.5pt single\">",
                                         ArtifactsDir + u"Table.AllowAutoFitOnTable.html");
        }
    }

    void KeepTableTogether()
    {
        //ExStart
        //ExFor:ParagraphFormat.KeepWithNext
        //ExFor:Row.IsLastRow
        //ExFor:Paragraph.IsEndOfCell
        //ExFor:Paragraph.IsInCell
        //ExFor:Cell.ParentRow
        //ExFor:Cell.Paragraphs
        //ExSummary:Shows how to set a table to stay together on the same page.
        auto doc = MakeObject<Document>(MyDir + u"Table spanning two pages.docx");
        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        // Enabling KeepWithNext for every paragraph in the table except for the
        // last ones in the last row will prevent the table from splitting across multiple pages.
        for (auto cell : System::IterateOver(table->GetChildNodes(NodeType::Cell, true)->LINQ_OfType<SharedPtr<Cell>>()))
        {
            for (auto para : System::IterateOver(cell->get_Paragraphs()->LINQ_OfType<SharedPtr<Paragraph>>()))
            {
                ASSERT_TRUE(para->get_IsInCell());

                if (!(cell->get_ParentRow()->get_IsLastRow() && para->get_IsEndOfCell()))
                {
                    para->get_ParagraphFormat()->set_KeepWithNext(true);
                }
            }
        }

        doc->Save(ArtifactsDir + u"Table.KeepTableTogether.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.KeepTableTogether.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        for (auto para : System::IterateOver(table->GetChildNodes(NodeType::Paragraph, true)->LINQ_OfType<SharedPtr<Paragraph>>()))
        {
            if (para->get_IsEndOfCell() && (System::DynamicCast<Cell>(para->get_ParentNode()))->get_ParentRow()->get_IsLastRow())
            {
                ASSERT_FALSE(para->get_ParagraphFormat()->get_KeepWithNext());
            }
            else
            {
                ASSERT_TRUE(para->get_ParagraphFormat()->get_KeepWithNext());
            }
        }
    }

    void GetIndexOfTableElements()
    {
        //ExStart
        //ExFor:NodeCollection.IndexOf(Node)
        //ExSummary:Shows how to get the index of a node in a collection.
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        SharedPtr<NodeCollection> allTables = doc->GetChildNodes(NodeType::Table, true);

        ASSERT_EQ(0, allTables->IndexOf(table));

        SharedPtr<Row> row = table->get_Rows()->idx_get(2);

        ASSERT_EQ(2, table->IndexOf(row));

        SharedPtr<Cell> cell = row->get_LastCell();

        ASSERT_EQ(4, row->IndexOf(cell));
        //ExEnd
    }

    void GetPreferredWidthTypeAndValue()
    {
        //ExStart
        //ExFor:PreferredWidthType
        //ExFor:PreferredWidth.Type
        //ExFor:PreferredWidth.Value
        //ExSummary:Shows how to verify the preferred width type and value of a table cell.
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        SharedPtr<Cell> firstCell = table->get_FirstRow()->get_FirstCell();

        ASSERT_EQ(PreferredWidthType::Percent, firstCell->get_CellFormat()->get_PreferredWidth()->get_Type());
        ASPOSE_ASSERT_EQ(11.16, firstCell->get_CellFormat()->get_PreferredWidth()->get_Value());
        //ExEnd
    }

    void AllowCellSpacing(bool allowCellSpacing)
    {
        //ExStart
        //ExFor:Table.AllowCellSpacing
        //ExFor:Table.CellSpacing
        //ExSummary:Shows how to enable spacing between individual cells in a table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Animal");
        builder->InsertCell();
        builder->Write(u"Class");
        builder->EndRow();
        builder->InsertCell();
        builder->Write(u"Dog");
        builder->InsertCell();
        builder->Write(u"Mammal");
        builder->EndTable();

        table->set_CellSpacing(3);

        // Set the "AllowCellSpacing" property to "true" to enable spacing between cells
        // with a magnitude equal to the value of the "CellSpacing" property, in points.
        // Set the "AllowCellSpacing" property to "false" to disable cell spacing
        // and ignore the value of the "CellSpacing" property.
        table->set_AllowCellSpacing(allowCellSpacing);

        doc->Save(ArtifactsDir + u"Table.AllowCellSpacing.html");

        // Adjusting the "CellSpacing" property will automatically enable cell spacing.
        table->set_CellSpacing(5);

        ASSERT_TRUE(table->get_AllowCellSpacing());
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.AllowCellSpacing.html");
        table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));

        ASPOSE_ASSERT_EQ(allowCellSpacing, table->get_AllowCellSpacing());

        if (allowCellSpacing)
        {
            ASPOSE_ASSERT_EQ(3.0, table->get_CellSpacing());
        }
        else
        {
            ASPOSE_ASSERT_EQ(0.0, table->get_CellSpacing());
        }

        TestUtil::FileContainsString(
            allowCellSpacing
                ? u"<td style=\"border-style:solid; border-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border:0.5pt "
                  u"single\">"
                : String(u"<td style=\"border-right-style:solid; border-right-width:0.75pt; border-bottom-style:solid; border-bottom-width:0.75pt; ") +
                      u"padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-bottom:0.5pt single; -aw-border-right:0.5pt single\">",
            ArtifactsDir + u"Table.AllowCellSpacing.html");
    }

    //ExStart
    //ExFor:Table
    //ExFor:Row
    //ExFor:Cell
    //ExFor:Table.#ctor(DocumentBase)
    //ExFor:Table.Title
    //ExFor:Table.Description
    //ExFor:Row.#ctor(DocumentBase)
    //ExFor:Cell.#ctor(DocumentBase)
    //ExFor:Cell.FirstParagraph
    //ExSummary:Shows how to build a nested table without using a document builder.
    void CreateNestedTable()
    {
        auto doc = MakeObject<Document>();

        // Create the outer table with three rows and four columns, and then add it to the document.
        SharedPtr<Table> outerTable = CreateTable(doc, 3, 4, u"Outer Table");
        doc->get_FirstSection()->get_Body()->AppendChild(outerTable);

        // Create another table with two rows and two columns and then insert it into the first table's first cell.
        SharedPtr<Table> innerTable = CreateTable(doc, 2, 2, u"Inner Table");
        outerTable->get_FirstRow()->get_FirstCell()->AppendChild(innerTable);

        doc->Save(ArtifactsDir + u"Table.CreateNestedTable.docx");
        TestCreateNestedTable(MakeObject<Document>(ArtifactsDir + u"Table.CreateNestedTable.docx"));
        //ExSkip
    }

    /// <summary>
    /// Creates a new table in the document with the given dimensions and text in each cell.
    /// </summary>
    static SharedPtr<Table> CreateTable(SharedPtr<Document> doc, int rowCount, int cellCount, String cellText)
    {
        auto table = MakeObject<Table>(doc);

        for (int rowId = 1; rowId <= rowCount; rowId++)
        {
            auto row = MakeObject<Row>(doc);
            table->AppendChild(row);

            for (int cellId = 1; cellId <= cellCount; cellId++)
            {
                auto cell = MakeObject<Cell>(doc);
                cell->AppendChild(MakeObject<Paragraph>(doc));
                cell->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, cellText));

                row->AppendChild(cell);
            }
        }

        // You can use the "Title" and "Description" properties to add a title and description respectively to your table.
        // The table must have at least one row before we can use these properties.
        // These properties are meaningful for ISO / IEC 29500 compliant .docx documents (see the OoxmlCompliance class).
        // If we save the document to pre-ISO/IEC 29500 formats, Microsoft Word ignores these properties.
        table->set_Title(u"Aspose table title");
        table->set_Description(u"Aspose table description");

        return table;
    }
    //ExEnd

    void TestCreateNestedTable(SharedPtr<Document> doc)
    {
        SharedPtr<Table> outerTable = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        auto innerTable = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 1, true));

        ASSERT_EQ(2, doc->GetChildNodes(NodeType::Table, true)->get_Count());
        ASSERT_EQ(1, outerTable->get_FirstRow()->get_FirstCell()->get_Tables()->get_Count());
        ASSERT_EQ(16, outerTable->GetChildNodes(NodeType::Cell, true)->get_Count());
        ASSERT_EQ(4, innerTable->GetChildNodes(NodeType::Cell, true)->get_Count());
        ASSERT_EQ(u"Aspose table title", innerTable->get_Title());
        ASSERT_EQ(u"Aspose table description", innerTable->get_Description());
    }

    //ExStart
    //ExFor:CellFormat.HorizontalMerge
    //ExFor:CellFormat.VerticalMerge
    //ExFor:CellMerge
    //ExSummary:Prints the horizontal and vertical merge type of a cell.
    void CheckCellsMerged()
    {
        auto doc = MakeObject<Document>(MyDir + u"Table with merged cells.docx");
        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        for (auto row : System::IterateOver(table->get_Rows()->LINQ_OfType<SharedPtr<Row>>()))
        {
            for (auto cell : System::IterateOver(row->get_Cells()->LINQ_OfType<SharedPtr<Cell>>()))
            {
                std::cout << PrintCellMergeType(cell) << std::endl;
            }
        }
        ASSERT_EQ(u"The cell at R1, C1 is vertically merged", PrintCellMergeType(table->get_FirstRow()->get_FirstCell()));
        //ExSkip
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

        return isVerticallyMerged ? String::Format(u"The cell at {0} is vertically merged", cellLocation)
                                  : String::Format(u"The cell at {0} is not merged", cellLocation);
    }
    //ExEnd

    void MergeCellRange()
    {
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        // We want to merge the range of cells found in between these two cells.
        SharedPtr<Cell> cellStartRange = table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2);
        SharedPtr<Cell> cellEndRange = table->get_Rows()->idx_get(3)->get_Cells()->idx_get(3);

        // Merge all the cells between the two specified cells into one.
        MergeCells(cellStartRange, cellEndRange);

        doc->Save(ArtifactsDir + u"Table.MergeCellRange.doc");

        int mergedCellsCount = 0;
        for (auto node : System::IterateOver(table->GetChildNodes(NodeType::Cell, true)))
        {
            auto cell = System::DynamicCast<Cell>(node);
            if (cell->get_CellFormat()->get_HorizontalMerge() != CellMerge::None || cell->get_CellFormat()->get_VerticalMerge() != CellMerge::None)
            {
                mergedCellsCount++;
            }
        }

        ASSERT_EQ(4, mergedCellsCount);
        ASSERT_TRUE(table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2)->get_CellFormat()->get_HorizontalMerge() == CellMerge::First);
        ASSERT_TRUE(table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2)->get_CellFormat()->get_VerticalMerge() == CellMerge::First);
        ASSERT_TRUE(table->get_Rows()->idx_get(3)->get_Cells()->idx_get(3)->get_CellFormat()->get_HorizontalMerge() == CellMerge::Previous);
        ASSERT_TRUE(table->get_Rows()->idx_get(3)->get_Cells()->idx_get(3)->get_CellFormat()->get_VerticalMerge() == CellMerge::Previous);
    }

    /// <summary>
    /// Merges the range of cells found between the two specified cells both horizontally and vertically.
    /// Can span over multiple rows.
    /// </summary>
    static void MergeCells(SharedPtr<Cell> startCell, SharedPtr<Cell> endCell)
    {
        SharedPtr<Table> parentTable = startCell->get_ParentRow()->get_ParentTable();

        // Find the row and cell indices for the start and end cells.
        System::Drawing::Point startCellPos(startCell->get_ParentRow()->IndexOf(startCell), parentTable->IndexOf(startCell->get_ParentRow()));
        System::Drawing::Point endCellPos(endCell->get_ParentRow()->IndexOf(endCell), parentTable->IndexOf(endCell->get_ParentRow()));

        // Create a range of cells to be merged based on these indices.
        // Inverse each index if the end cell is before the start cell.
        System::Drawing::Rectangle mergeRange(
            System::Math::Min(startCellPos.get_X(), endCellPos.get_X()), System::Math::Min(startCellPos.get_Y(), endCellPos.get_Y()),
            System::Math::Abs(endCellPos.get_X() - startCellPos.get_X()) + 1, System::Math::Abs(endCellPos.get_Y() - startCellPos.get_Y()) + 1);

        for (auto row : System::IterateOver(parentTable->get_Rows()->LINQ_OfType<SharedPtr<Row>>()))
        {
            for (auto cell : System::IterateOver(row->get_Cells()->LINQ_OfType<SharedPtr<Cell>>()))
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

    void CombineTables()
    {
        //ExStart
        //ExFor:Cell.CellFormat
        //ExFor:CellFormat.Borders
        //ExFor:Table.Rows
        //ExFor:Table.FirstRow
        //ExFor:CellFormat.ClearFormatting
        //ExFor:CompositeNode.HasChildNodes
        //ExSummary:Shows how to combine the rows from two tables into one.
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        // Below are two ways of getting a table from a document.
        // 1 -  From the "Tables" collection of a Body node:
        SharedPtr<Table> firstTable = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        // 2 -  Using the "GetChild" method:
        auto secondTable = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 1, true));

        // Append all rows from the current table to the next.
        while (secondTable->get_HasChildNodes())
        {
            firstTable->get_Rows()->Add(secondTable->get_FirstRow());
        }

        // Remove the empty table container.
        secondTable->Remove();

        doc->Save(ArtifactsDir + u"Table.CombineTables.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.CombineTables.docx");

        ASSERT_EQ(1, doc->GetChildNodes(NodeType::Table, true)->get_Count());
        ASSERT_EQ(9, doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_Rows()->get_Count());
        ASSERT_EQ(42, doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->GetChildNodes(NodeType::Cell, true)->get_Count());
    }

    void SplitTable()
    {
        auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

        SharedPtr<Table> firstTable = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

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

        doc = DocumentHelper::SaveOpen(doc);

        ASPOSE_ASSERT_EQ(row, table->get_FirstRow());
        ASSERT_EQ(2, firstTable->get_Rows()->get_Count());
        ASSERT_EQ(3, table->get_Rows()->get_Count());
        ASSERT_EQ(3, doc->GetChildNodes(NodeType::Table, true)->get_Count());
    }

    void WrapText()
    {
        //ExStart
        //ExFor:Table.TextWrapping
        //ExFor:TextWrapping
        //ExSummary:Shows how to work with table text wrapping.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Cell 1");
        builder->InsertCell();
        builder->Write(u"Cell 2");
        builder->EndTable();
        table->set_PreferredWidth(PreferredWidth::FromPoints(300));

        builder->get_Font()->set_Size(16);
        builder->Writeln(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Set the "TextWrapping" property to "TextWrapping.Around" to get the table to wrap text around it,
        // and push it down into the paragraph below by setting the position.
        table->set_TextWrapping(TextWrapping::Around);
        table->set_AbsoluteHorizontalDistance(100);
        table->set_AbsoluteVerticalDistance(20);

        doc->Save(ArtifactsDir + u"Table.WrapText.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.WrapText.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(TextWrapping::Around, table->get_TextWrapping());
        ASPOSE_ASSERT_EQ(100.0, table->get_AbsoluteHorizontalDistance());
        ASPOSE_ASSERT_EQ(20.0, table->get_AbsoluteVerticalDistance());
    }

    void GetFloatingTableProperties()
    {
        //ExStart
        //ExFor:Table.HorizontalAnchor
        //ExFor:Table.VerticalAnchor
        //ExFor:Table.AllowOverlap
        //ExFor:ShapeBase.AllowOverlap
        //ExSummary:Shows how to work with floating tables properties.
        auto doc = MakeObject<Document>(MyDir + u"Table wrapped by text.docx");

        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        if (table->get_TextWrapping() == TextWrapping::Around)
        {
            ASSERT_EQ(RelativeHorizontalPosition::Margin, table->get_HorizontalAnchor());
            ASSERT_EQ(RelativeVerticalPosition::Paragraph, table->get_VerticalAnchor());
            ASPOSE_ASSERT_EQ(false, table->get_AllowOverlap());

            // Only Margin, Page, Column available in RelativeHorizontalPosition for HorizontalAnchor setter.
            // The ArgumentException will be thrown for any other values.
            table->set_HorizontalAnchor(RelativeHorizontalPosition::Column);

            // Only Margin, Page, Paragraph available in RelativeVerticalPosition for VerticalAnchor setter.
            // The ArgumentException will be thrown for any other values.
            table->set_VerticalAnchor(RelativeVerticalPosition::Page);
        }
        //ExEnd
    }

    void ChangeFloatingTableProperties()
    {
        //ExStart
        //ExFor:Table.RelativeHorizontalAlignment
        //ExFor:Table.RelativeVerticalAlignment
        //ExFor:Table.AbsoluteHorizontalDistance
        //ExFor:Table.AbsoluteVerticalDistance
        //ExSummary:Shows how set the location of floating tables.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Table 1, cell 1");
        builder->EndTable();
        table->set_PreferredWidth(PreferredWidth::FromPoints(300));

        // Set the table's location to a place on the page, such as, in this case, the bottom right corner.
        table->set_RelativeVerticalAlignment(VerticalAlignment::Bottom);
        table->set_RelativeHorizontalAlignment(HorizontalAlignment::Right);

        table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Table 2, cell 1");
        builder->EndTable();
        table->set_PreferredWidth(PreferredWidth::FromPoints(300));

        // We can also set a horizontal and vertical offset in points from the paragraph's location where we inserted the table.
        table->set_AbsoluteVerticalDistance(50);
        table->set_AbsoluteHorizontalDistance(100);

        doc->Save(ArtifactsDir + u"Table.ChangeFloatingTableProperties.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.ChangeFloatingTableProperties.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(VerticalAlignment::Bottom, table->get_RelativeVerticalAlignment());
        ASSERT_EQ(HorizontalAlignment::Right, table->get_RelativeHorizontalAlignment());

        table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 1, true));

        ASPOSE_ASSERT_EQ(50.0, table->get_AbsoluteVerticalDistance());
        ASPOSE_ASSERT_EQ(100.0, table->get_AbsoluteHorizontalDistance());
    }

    void TableStyleCreation()
    {
        //ExStart
        //ExFor:Table.Bidi
        //ExFor:Table.CellSpacing
        //ExFor:Table.Style
        //ExFor:Table.StyleName
        //ExFor:TableStyle
        //ExFor:TableStyle.AllowBreakAcrossPages
        //ExFor:TableStyle.Bidi
        //ExFor:TableStyle.CellSpacing
        //ExFor:TableStyle.BottomPadding
        //ExFor:TableStyle.LeftPadding
        //ExFor:TableStyle.RightPadding
        //ExFor:TableStyle.TopPadding
        //ExFor:TableStyle.Shading
        //ExFor:TableStyle.Borders
        //ExFor:TableStyle.VerticalAlignment
        //ExSummary:Shows how to create custom style settings for the table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Name");
        builder->InsertCell();
        builder->Write(u"مرحبًا");
        builder->EndRow();
        builder->InsertCell();
        builder->InsertCell();
        builder->EndTable();

        auto tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->Add(StyleType::Table, u"MyTableStyle1"));
        tableStyle->set_AllowBreakAcrossPages(true);
        tableStyle->set_Bidi(true);
        tableStyle->set_CellSpacing(5);
        tableStyle->set_BottomPadding(20);
        tableStyle->set_LeftPadding(5);
        tableStyle->set_RightPadding(10);
        tableStyle->set_TopPadding(20);
        tableStyle->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_AntiqueWhite());
        tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Blue());
        tableStyle->get_Borders()->set_LineStyle(LineStyle::DotDash);
        tableStyle->set_VerticalAlignment(CellVerticalAlignment::Center);

        table->set_Style(tableStyle);

        // Setting the style properties of a table may affect the properties of the table itself.
        ASSERT_TRUE(table->get_Bidi());
        ASPOSE_ASSERT_EQ(5.0, table->get_CellSpacing());
        ASSERT_EQ(u"MyTableStyle1", table->get_StyleName());

        doc->Save(ArtifactsDir + u"Table.TableStyleCreation.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.TableStyleCreation.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_TRUE(table->get_Bidi());
        ASPOSE_ASSERT_EQ(5.0, table->get_CellSpacing());
        ASSERT_EQ(u"MyTableStyle1", table->get_StyleName());
        ASPOSE_ASSERT_EQ(20.0, tableStyle->get_BottomPadding());
        ASPOSE_ASSERT_EQ(5.0, tableStyle->get_LeftPadding());
        ASPOSE_ASSERT_EQ(10.0, tableStyle->get_RightPadding());
        ASPOSE_ASSERT_EQ(20.0, tableStyle->get_TopPadding());
        ASSERT_EQ(6, table->get_FirstRow()->get_RowFormat()->get_Borders()->LINQ_Count(
                         static_cast<System::Func<SharedPtr<Border>, bool>>(static_cast<std::function<bool(SharedPtr<Border> b)>>(
                             [](SharedPtr<Border> b) -> bool { return b->get_Color().ToArgb() == System::Drawing::Color::get_Blue().ToArgb(); }))));
        ASSERT_EQ(CellVerticalAlignment::Center, tableStyle->get_VerticalAlignment());

        tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle1"));

        ASSERT_TRUE(tableStyle->get_AllowBreakAcrossPages());
        ASSERT_TRUE(tableStyle->get_Bidi());
        ASPOSE_ASSERT_EQ(5.0, tableStyle->get_CellSpacing());
        ASPOSE_ASSERT_EQ(20.0, tableStyle->get_BottomPadding());
        ASPOSE_ASSERT_EQ(5.0, tableStyle->get_LeftPadding());
        ASPOSE_ASSERT_EQ(10.0, tableStyle->get_RightPadding());
        ASPOSE_ASSERT_EQ(20.0, tableStyle->get_TopPadding());
        ASSERT_EQ(System::Drawing::Color::get_AntiqueWhite().ToArgb(), tableStyle->get_Shading()->get_BackgroundPatternColor().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), tableStyle->get_Borders()->get_Color().ToArgb());
        ASSERT_EQ(LineStyle::DotDash, tableStyle->get_Borders()->get_LineStyle());
        ASSERT_EQ(CellVerticalAlignment::Center, tableStyle->get_VerticalAlignment());
    }

    void SetTableAlignment()
    {
        //ExStart
        //ExFor:TableStyle.Alignment
        //ExFor:TableStyle.LeftIndent
        //ExSummary:Shows how to set the position of a table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Below are two ways of aligning a table horizontally.
        // 1 -  Use the "Alignment" property to align it to a location on the page, such as the center:
        auto tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->Add(StyleType::Table, u"MyTableStyle1"));
        tableStyle->set_Alignment(TableAlignment::Center);
        tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Blue());
        tableStyle->get_Borders()->set_LineStyle(LineStyle::Single);

        // Insert a table and apply the style we created to it.
        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Aligned to the center of the page");
        builder->EndTable();
        table->set_PreferredWidth(PreferredWidth::FromPoints(300));

        table->set_Style(tableStyle);

        // 2 -  Use the "LeftIndent" to specify an indent from the left margin of the page:
        tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->Add(StyleType::Table, u"MyTableStyle2"));
        tableStyle->set_LeftIndent(55);
        tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Green());
        tableStyle->get_Borders()->set_LineStyle(LineStyle::Single);

        table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Aligned according to left indent");
        builder->EndTable();
        table->set_PreferredWidth(PreferredWidth::FromPoints(300));

        table->set_Style(tableStyle);

        doc->Save(ArtifactsDir + u"Table.SetTableAlignment.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.SetTableAlignment.docx");

        tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle1"));

        ASSERT_EQ(TableAlignment::Center, tableStyle->get_Alignment());
        ASPOSE_ASSERT_EQ(tableStyle, doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_Style());

        tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle2"));

        ASPOSE_ASSERT_EQ(55.0, tableStyle->get_LeftIndent());
        ASPOSE_ASSERT_EQ(tableStyle, (System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 1, true)))->get_Style());
    }

    void ConditionalStyles()
    {
        //ExStart
        //ExFor:ConditionalStyle
        //ExFor:ConditionalStyle.Shading
        //ExFor:ConditionalStyle.Borders
        //ExFor:ConditionalStyle.ParagraphFormat
        //ExFor:ConditionalStyle.BottomPadding
        //ExFor:ConditionalStyle.LeftPadding
        //ExFor:ConditionalStyle.RightPadding
        //ExFor:ConditionalStyle.TopPadding
        //ExFor:ConditionalStyle.Font
        //ExFor:ConditionalStyle.Type
        //ExFor:ConditionalStyleCollection.GetEnumerator
        //ExFor:ConditionalStyleCollection.FirstRow
        //ExFor:ConditionalStyleCollection.LastRow
        //ExFor:ConditionalStyleCollection.LastColumn
        //ExFor:ConditionalStyleCollection.Count
        //ExFor:ConditionalStyleCollection
        //ExFor:ConditionalStyleCollection.BottomLeftCell
        //ExFor:ConditionalStyleCollection.BottomRightCell
        //ExFor:ConditionalStyleCollection.EvenColumnBanding
        //ExFor:ConditionalStyleCollection.EvenRowBanding
        //ExFor:ConditionalStyleCollection.FirstColumn
        //ExFor:ConditionalStyleCollection.Item(ConditionalStyleType)
        //ExFor:ConditionalStyleCollection.Item(Int32)
        //ExFor:ConditionalStyleCollection.OddColumnBanding
        //ExFor:ConditionalStyleCollection.OddRowBanding
        //ExFor:ConditionalStyleCollection.TopLeftCell
        //ExFor:ConditionalStyleCollection.TopRightCell
        //ExFor:ConditionalStyleType
        //ExFor:TableStyle.ConditionalStyles
        //ExSummary:Shows how to work with certain area styles of a table.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Cell 1");
        builder->InsertCell();
        builder->Write(u"Cell 2");
        builder->EndRow();
        builder->InsertCell();
        builder->Write(u"Cell 3");
        builder->InsertCell();
        builder->Write(u"Cell 4");
        builder->EndTable();

        // Create a custom table style.
        auto tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->Add(StyleType::Table, u"MyTableStyle1"));

        // Conditional styles are formatting changes that affect only some of the table's cells
        // based on a predicate, such as the cells being in the last row.
        // Below are three ways of accessing a table style's conditional styles from the "ConditionalStyles" collection.
        // 1 -  By style type:
        tableStyle->get_ConditionalStyles()
            ->idx_get(ConditionalStyleType::FirstRow)
            ->get_Shading()
            ->set_BackgroundPatternColor(System::Drawing::Color::get_AliceBlue());

        // 2 -  By index:
        tableStyle->get_ConditionalStyles()->idx_get(0)->get_Borders()->set_Color(System::Drawing::Color::get_Black());
        tableStyle->get_ConditionalStyles()->idx_get(0)->get_Borders()->set_LineStyle(LineStyle::DotDash);
        ASSERT_EQ(ConditionalStyleType::FirstRow, tableStyle->get_ConditionalStyles()->idx_get(0)->get_Type());

        // 3 -  As a property:
        tableStyle->get_ConditionalStyles()->get_FirstRow()->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);

        // Apply padding and text formatting to conditional styles.
        tableStyle->get_ConditionalStyles()->get_LastRow()->set_BottomPadding(10);
        tableStyle->get_ConditionalStyles()->get_LastRow()->set_LeftPadding(10);
        tableStyle->get_ConditionalStyles()->get_LastRow()->set_RightPadding(10);
        tableStyle->get_ConditionalStyles()->get_LastRow()->set_TopPadding(10);
        tableStyle->get_ConditionalStyles()->get_LastColumn()->get_Font()->set_Bold(true);

        // List all possible style conditions.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<ConditionalStyle>>> enumerator = tableStyle->get_ConditionalStyles()->GetEnumerator();
            while (enumerator->MoveNext())
            {
                SharedPtr<ConditionalStyle> currentStyle = enumerator->get_Current();
                if (currentStyle != nullptr)
                {
                    std::cout << System::EnumGetName(currentStyle->get_Type()) << std::endl;
                }
            }
        }

        // Apply the custom style, which contains all conditional styles, to the table.
        table->set_Style(tableStyle);

        // Our style applies some conditional styles by default.
        ASSERT_EQ(TableStyleOptions::FirstRow | TableStyleOptions::FirstColumn | TableStyleOptions::RowBands, table->get_StyleOptions());

        // We will need to enable all other styles ourselves via the "StyleOptions" property.
        table->set_StyleOptions(table->get_StyleOptions() | TableStyleOptions::LastRow | TableStyleOptions::LastColumn);

        doc->Save(ArtifactsDir + u"Table.ConditionalStyles.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.ConditionalStyles.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        ASSERT_EQ(TableStyleOptions::Default | TableStyleOptions::LastRow | TableStyleOptions::LastColumn, table->get_StyleOptions());
        SharedPtr<ConditionalStyleCollection> conditionalStyles =
            (System::DynamicCast<TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle1")))->get_ConditionalStyles();

        ASSERT_EQ(ConditionalStyleType::FirstRow, conditionalStyles->idx_get(0)->get_Type());
        ASSERT_EQ(System::Drawing::Color::get_AliceBlue().ToArgb(), conditionalStyles->idx_get(0)->get_Shading()->get_BackgroundPatternColor().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_Black().ToArgb(), conditionalStyles->idx_get(0)->get_Borders()->get_Color().ToArgb());
        ASSERT_EQ(LineStyle::DotDash, conditionalStyles->idx_get(0)->get_Borders()->get_LineStyle());
        ASSERT_EQ(ParagraphAlignment::Center, conditionalStyles->idx_get(0)->get_ParagraphFormat()->get_Alignment());

        ASSERT_EQ(ConditionalStyleType::LastRow, conditionalStyles->idx_get(2)->get_Type());
        ASPOSE_ASSERT_EQ(10.0, conditionalStyles->idx_get(2)->get_BottomPadding());
        ASPOSE_ASSERT_EQ(10.0, conditionalStyles->idx_get(2)->get_LeftPadding());
        ASPOSE_ASSERT_EQ(10.0, conditionalStyles->idx_get(2)->get_RightPadding());
        ASPOSE_ASSERT_EQ(10.0, conditionalStyles->idx_get(2)->get_TopPadding());

        ASSERT_EQ(ConditionalStyleType::LastColumn, conditionalStyles->idx_get(3)->get_Type());
        ASSERT_TRUE(conditionalStyles->idx_get(3)->get_Font()->get_Bold());
    }

    void ClearTableStyleFormatting()
    {
        //ExStart
        //ExFor:ConditionalStyle.ClearFormatting
        //ExFor:ConditionalStyleCollection.ClearFormatting
        //ExSummary:Shows how to reset conditional table styles.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"First row");
        builder->EndRow();
        builder->InsertCell();
        builder->Write(u"Last row");
        builder->EndTable();

        auto tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->Add(StyleType::Table, u"MyTableStyle1"));
        table->set_Style(tableStyle);

        // Set the table style to color the borders of the first row of the table in red.
        tableStyle->get_ConditionalStyles()->get_FirstRow()->get_Borders()->set_Color(System::Drawing::Color::get_Red());

        // Set the table style to color the borders of the last row of the table in blue.
        tableStyle->get_ConditionalStyles()->get_LastRow()->get_Borders()->set_Color(System::Drawing::Color::get_Blue());

        // Below are two ways of using the "ClearFormatting" method to clear the conditional styles.
        // 1 -  Clear the conditional styles for a specific part of a table:
        tableStyle->get_ConditionalStyles()->idx_get(0)->ClearFormatting();

        ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty, tableStyle->get_ConditionalStyles()->get_FirstRow()->get_Borders()->get_Color());

        // 2 -  Clear the conditional styles for the entire table:
        tableStyle->get_ConditionalStyles()->ClearFormatting();

        auto isEmptyColor = [](SharedPtr<ConditionalStyle> s)
        {
            return s->get_Borders()->get_Color() == System::Drawing::Color::Empty;
        };
        ASSERT_TRUE(tableStyle->get_ConditionalStyles()->LINQ_All(isEmptyColor));
        //ExEnd
    }

    void AlternatingRowStyles()
    {
        //ExStart
        //ExFor:TableStyle.ColumnStripe
        //ExFor:TableStyle.RowStripe
        //ExSummary:Shows how to create conditional table styles that alternate between rows.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // We can configure a conditional style of a table to apply a different color to the row/column,
        // based on whether the row/column is even or odd, creating an alternating color pattern.
        // We can also apply a number n to the row/column banding,
        // meaning that the color alternates after every n rows/columns instead of one.
        // Create a table where single columns and rows will band the columns will banded in threes.
        SharedPtr<Table> table = builder->StartTable();
        for (int i = 0; i < 15; i++)
        {
            for (int j = 0; j < 4; j++)
            {
                builder->InsertCell();
                builder->Writeln(String::Format(u"{0} column.", (j % 2 == 0 ? String(u"Even") : String(u"Odd"))));
                builder->Write(String::Format(u"Row banding {0}.", (i % 3 == 0 ? String(u"start") : String(u"continuation"))));
            }
            builder->EndRow();
        }
        builder->EndTable();

        // Apply a line style to all the borders of the table.
        auto tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->Add(StyleType::Table, u"MyTableStyle1"));
        tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Black());
        tableStyle->get_Borders()->set_LineStyle(LineStyle::Double);

        // Set the two colors, which will alternate over every 3 rows.
        tableStyle->set_RowStripe(3);
        tableStyle->get_ConditionalStyles()
            ->idx_get(ConditionalStyleType::OddRowBanding)
            ->get_Shading()
            ->set_BackgroundPatternColor(System::Drawing::Color::get_LightBlue());
        tableStyle->get_ConditionalStyles()
            ->idx_get(ConditionalStyleType::EvenRowBanding)
            ->get_Shading()
            ->set_BackgroundPatternColor(System::Drawing::Color::get_LightCyan());

        // Set a color to apply to every even column, which will override any custom row coloring.
        tableStyle->set_ColumnStripe(1);
        tableStyle->get_ConditionalStyles()
            ->idx_get(ConditionalStyleType::EvenColumnBanding)
            ->get_Shading()
            ->set_BackgroundPatternColor(System::Drawing::Color::get_LightSalmon());

        table->set_Style(tableStyle);

        // The "StyleOptions" property enables row banding by default.
        ASSERT_EQ(TableStyleOptions::FirstRow | TableStyleOptions::FirstColumn | TableStyleOptions::RowBands, table->get_StyleOptions());

        // Use the "StyleOptions" property also to enable column banding.
        table->set_StyleOptions(table->get_StyleOptions() | TableStyleOptions::ColumnBands);

        doc->Save(ArtifactsDir + u"Table.AlternatingRowStyles.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Table.AlternatingRowStyles.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle1"));

        ASPOSE_ASSERT_EQ(tableStyle, table->get_Style());
        ASSERT_EQ(table->get_StyleOptions() | TableStyleOptions::ColumnBands, table->get_StyleOptions());

        ASSERT_EQ(System::Drawing::Color::get_Black().ToArgb(), tableStyle->get_Borders()->get_Color().ToArgb());
        ASSERT_EQ(LineStyle::Double, tableStyle->get_Borders()->get_LineStyle());
        ASSERT_EQ(3, tableStyle->get_RowStripe());
        ASSERT_EQ(System::Drawing::Color::get_LightBlue().ToArgb(),
                  tableStyle->get_ConditionalStyles()->idx_get(ConditionalStyleType::OddRowBanding)->get_Shading()->get_BackgroundPatternColor().ToArgb());
        ASSERT_EQ(System::Drawing::Color::get_LightCyan().ToArgb(),
                  tableStyle->get_ConditionalStyles()->idx_get(ConditionalStyleType::EvenRowBanding)->get_Shading()->get_BackgroundPatternColor().ToArgb());
        ASSERT_EQ(1, tableStyle->get_ColumnStripe());
        ASSERT_EQ(System::Drawing::Color::get_LightSalmon().ToArgb(),
                  tableStyle->get_ConditionalStyles()->idx_get(ConditionalStyleType::EvenColumnBanding)->get_Shading()->get_BackgroundPatternColor().ToArgb());
    }

    void ConvertToHorizontallyMergedCells()
    {
        //ExStart
        //ExFor:Table.ConvertToHorizontallyMergedCells
        //ExSummary:Shows how to convert cells horizontally merged by width to cells merged by CellFormat.HorizontalMerge.
        auto doc = MakeObject<Document>(MyDir + u"Table with merged cells.docx");

        // Microsoft Word does not write merge flags anymore, defining merged cells by width instead.
        // Aspose.Words by default define only 5 cells in a row, and none of them have the horizontal merge flag,
        // even though there were 7 cells in the row before the horizontal merging took place.
        SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
        SharedPtr<Row> row = table->get_Rows()->idx_get(0);

        ASSERT_EQ(5, row->get_Cells()->get_Count());
        auto horizontalMergeIsNone = [](SharedPtr<Node> c)
        {
            return (System::DynamicCast<Cell>(c))->get_CellFormat()->get_HorizontalMerge() == CellMerge::None;
        };
        ASSERT_TRUE(row->get_Cells()->LINQ_All(horizontalMergeIsNone));

        // Use the "ConvertToHorizontallyMergedCells" method to convert cells horizontally merged
        // by its width to the cell horizontally merged by flags.
        // Now, we have 7 cells, and some of them have horizontal merge values.
        table->ConvertToHorizontallyMergedCells();
        row = table->get_Rows()->idx_get(0);

        ASSERT_EQ(7, row->get_Cells()->get_Count());

        ASSERT_EQ(CellMerge::None, row->get_Cells()->idx_get(0)->get_CellFormat()->get_HorizontalMerge());
        ASSERT_EQ(CellMerge::First, row->get_Cells()->idx_get(1)->get_CellFormat()->get_HorizontalMerge());
        ASSERT_EQ(CellMerge::Previous, row->get_Cells()->idx_get(2)->get_CellFormat()->get_HorizontalMerge());
        ASSERT_EQ(CellMerge::None, row->get_Cells()->idx_get(3)->get_CellFormat()->get_HorizontalMerge());
        ASSERT_EQ(CellMerge::First, row->get_Cells()->idx_get(4)->get_CellFormat()->get_HorizontalMerge());
        ASSERT_EQ(CellMerge::Previous, row->get_Cells()->idx_get(5)->get_CellFormat()->get_HorizontalMerge());
        ASSERT_EQ(CellMerge::None, row->get_Cells()->idx_get(6)->get_CellFormat()->get_HorizontalMerge());
        //ExEnd
    }
};

} // namespace ApiExamples
