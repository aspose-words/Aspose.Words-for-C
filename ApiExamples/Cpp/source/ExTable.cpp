﻿// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTable.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/math.h>
#include <system/linq/enumerable.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/enum_helpers.h>
#include <system/details/dispose_guard.h>
#include <system/console.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/rectangle.h>
#include <drawing/point.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/TextOrientation.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/HeightRule.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Tables/TextWrapping.h>
#include <Aspose.Words.Cpp/Model/Tables/TableStyleOptions.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/TableAlignment.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidthType.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Model/Tables/CellMerge.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/AutoFitBehavior.h>
#include <Aspose.Words.Cpp/Model/Styles/TableStyle.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/ConditionalStyleType.h>
#include <Aspose.Words.Cpp/Model/Styles/ConditionalStyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/ConditionalStyle.h>
#include <Aspose.Words.Cpp/Model/Sections/StoryType.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Drawing/VerticalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Borders/TextureIndex.h>
#include <Aspose.Words.Cpp/Model/Borders/Shading.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderType.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>

#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Tables;
namespace ApiExamples {

int ExTable::GetNestedDepthOfTable(SharedPtr<Aspose::Words::Tables::Table> table)
{
    int depth = 0;

    // The parent of the table will be a Cell, instead attempt to find a grandparent that is of type Table
    SharedPtr<Node> parent = table->GetAncestor(table->get_NodeType());

    while (parent != nullptr)
    {
        // Every time we find a table a level up we increase the depth counter and then try to find an
        // ancestor of type table from the parent
        depth++;
        parent = parent->GetAncestorOf<SharedPtr<Table>>();
    }

    return depth;
}

int ExTable::GetChildTableCount(SharedPtr<Aspose::Words::Tables::Table> table)
{
    int tableCount = 0;
    // Iterate through all child rows in the table
    for (auto row : System::IterateOver(table->get_Rows()->LINQ_OfType<SharedPtr<Row> >()))
    {
        // Iterate through all child cells in the row
        for (auto Cell : System::IterateOver(row->get_Cells()->LINQ_OfType<SharedPtr<Cell> >()))
        {
            // Retrieve the collection of child tables of this cell
            SharedPtr<TableCollection> childTables = Cell->get_Tables();

            // If this cell has a table as a child then return true
            if (childTables->get_Count() > 0)
            {
                tableCount++;
            }
        }
    }

    // No cell contains a table
    return tableCount;
}

void ExTable::ConvertTextboxToTable(SharedPtr<Aspose::Words::Drawing::Shape> textBox)
{
    if (textBox->get_StoryType() != Aspose::Words::StoryType::Textbox)
    {
        throw System::ArgumentException(u"Can only convert a shape of type textbox");
    }

    auto doc = System::DynamicCast<Aspose::Words::Document>(textBox->get_Document());
    auto section = System::DynamicCast<Aspose::Words::Section>(textBox->GetAncestor(Aspose::Words::NodeType::Section));

    // Create a table to replace the textbox and transfer the same content and formatting
    auto table = MakeObject<Table>(doc);
    // Ensure that the table contains a row and a cell
    table->EnsureMinimum();
    // Use fixed column widths
    table->AutoFit(Aspose::Words::Tables::AutoFitBehavior::FixedColumnWidths);

    // A shape is inline level (within a paragraph) where a table can only be block level so insert the table
    // after the paragraph which contains the shape
    SharedPtr<Node> shapeParent = textBox->get_ParentNode();
    shapeParent->get_ParentNode()->InsertAfter(table, shapeParent);

    // If the textbox is not inline then try to match the shape's left position using the table's left indent
    if (!textBox->get_IsInline() && textBox->get_Left() < section->get_PageSetup()->get_PageWidth())
    {
        table->set_LeftIndent(textBox->get_Left());
    }

    // We are only using one cell to replicate a textbox so we can make use of the FirstRow and FirstCell property
    // Carry over borders and shading
    SharedPtr<Row> firstRow = table->get_FirstRow();
    SharedPtr<Cell> firstCell = firstRow->get_FirstCell();
    firstCell->get_CellFormat()->get_Borders()->set_Color(textBox->get_StrokeColor());
    firstCell->get_CellFormat()->get_Borders()->set_LineWidth(textBox->get_StrokeWeight());
    firstCell->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(textBox->get_Fill()->get_Color());

    // Transfer the same height and width of the textbox to the table
    firstRow->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Exactly);
    firstRow->get_RowFormat()->set_Height(textBox->get_Height());
    firstCell->get_CellFormat()->set_Width(textBox->get_Width());
    table->set_AllowAutoFit(false);

    // Replicate the textbox's horizontal alignment
    TableAlignment horizontalAlignment;
    switch (textBox->get_HorizontalAlignment())
    {
        case Aspose::Words::Drawing::HorizontalAlignment::Left:
            horizontalAlignment = Aspose::Words::Tables::TableAlignment::Left;
            break;

        case Aspose::Words::Drawing::HorizontalAlignment::Center:
            horizontalAlignment = Aspose::Words::Tables::TableAlignment::Center;
            break;

        case Aspose::Words::Drawing::HorizontalAlignment::Right:
            horizontalAlignment = Aspose::Words::Tables::TableAlignment::Right;
            break;

        default:
            horizontalAlignment = Aspose::Words::Tables::TableAlignment::Left;
            break;

    }

    table->set_Alignment(horizontalAlignment);
    firstCell->RemoveAllChildren();

    // Append all content from the textbox to the new table
    for (SharedPtr<Node> node : textBox->GetChildNodes(Aspose::Words::NodeType::Any, false)->ToArray())
    {
        table->get_FirstRow()->get_FirstCell()->AppendChild(node);
    }

    // Remove the empty textbox from the document
    textBox->Remove();
}

SharedPtr<Aspose::Words::Tables::Table> ExTable::CreateTable(SharedPtr<Aspose::Words::Document> doc, int rowCount, int cellCount, String cellText)
{
    auto table = MakeObject<Table>(doc);

    // Create the specified number of rows
    for (int rowId = 1; rowId <= rowCount; rowId++)
    {
        auto row = MakeObject<Row>(doc);
        table->AppendChild(row);

        // Create the specified number of cells for each row
        for (int cellId = 1; cellId <= cellCount; cellId++)
        {
            auto cell = MakeObject<Cell>(doc);
            row->AppendChild(cell);
            // Add a blank paragraph to the cell
            cell->AppendChild(MakeObject<Paragraph>(doc));

            // Add the text
            cell->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, cellText));
        }
    }

    // You can add title and description to your table only when added at least one row to the table first
    // This properties are meaningful for ISO / IEC 29500 compliant DOCX documents(see the OoxmlCompliance class)
    // When saved to pre-ISO/IEC 29500 formats, the properties are ignored
    table->set_Title(u"Aspose table title");
    table->set_Description(u"Aspose table description");

    return table;
}

void ExTable::TestCreateNestedTable(SharedPtr<Aspose::Words::Document> doc)
{
    auto outerTable = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    auto innerTable = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));

    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
    ASSERT_EQ(1, outerTable->get_FirstRow()->get_FirstCell()->get_Tables()->get_Count());
    ASSERT_EQ(16, outerTable->GetChildNodes(Aspose::Words::NodeType::Cell, true)->get_Count());
    ASSERT_EQ(4, innerTable->GetChildNodes(Aspose::Words::NodeType::Cell, true)->get_Count());
    ASSERT_EQ(u"Aspose table title", innerTable->get_Title());
    ASSERT_EQ(u"Aspose table description", innerTable->get_Description());
}

String ExTable::PrintCellMergeType(SharedPtr<Aspose::Words::Tables::Cell> cell)
{
    bool isHorizontallyMerged = cell->get_CellFormat()->get_HorizontalMerge() != Aspose::Words::Tables::CellMerge::None;
    bool isVerticallyMerged = cell->get_CellFormat()->get_VerticalMerge() != Aspose::Words::Tables::CellMerge::None;
    String cellLocation = String::Format(u"R{0}, C{1}",cell->get_ParentRow()->get_ParentTable()->IndexOf(cell->get_ParentRow()) + 1,cell->get_ParentRow()->IndexOf(cell) + 1);

    if (isHorizontallyMerged && isVerticallyMerged)
    {
        return String::Format(u"The cell at {0} is both horizontally and vertically merged",cellLocation);
    }
    if (isHorizontallyMerged)
    {
        return String::Format(u"The cell at {0} is horizontally merged.",cellLocation);
    }

    return isVerticallyMerged ? String::Format(u"The cell at {0} is vertically merged",cellLocation) : String::Format(u"The cell at {0} is not merged",cellLocation);
}

void ExTable::MergeCells(SharedPtr<Aspose::Words::Tables::Cell> startCell, SharedPtr<Aspose::Words::Tables::Cell> endCell)
{
    SharedPtr<Table> parentTable = startCell->get_ParentRow()->get_ParentTable();

    // Find the row and cell indices for the start and end cell
    System::Drawing::Point startCellPos(startCell->get_ParentRow()->IndexOf(startCell), parentTable->IndexOf(startCell->get_ParentRow()));
    System::Drawing::Point endCellPos(endCell->get_ParentRow()->IndexOf(endCell), parentTable->IndexOf(endCell->get_ParentRow()));
    // Create the range of cells to be merged based off these indices
    // Inverse each index if the end cell if before the start cell
    System::Drawing::Rectangle mergeRange(System::Math::Min(startCellPos.get_X(), endCellPos.get_X()), System::Math::Min(startCellPos.get_Y(), endCellPos.get_Y()), System::Math::Abs(endCellPos.get_X() - startCellPos.get_X()) + 1, System::Math::Abs(endCellPos.get_Y() - startCellPos.get_Y()) + 1);

    for (auto row : System::IterateOver(parentTable->get_Rows()->LINQ_OfType<SharedPtr<Row> >()))
    {
        for (auto cell : System::IterateOver(row->get_Cells()->LINQ_OfType<SharedPtr<Cell> >()))
        {
            System::Drawing::Point currentPos(row->IndexOf(cell), parentTable->IndexOf(row));
            // Check if the current cell is inside our merge range then merge it
            if (mergeRange.Contains(currentPos))
            {
                cell->get_CellFormat()->set_HorizontalMerge(currentPos.get_X() == mergeRange.get_X() ? Aspose::Words::Tables::CellMerge::First : Aspose::Words::Tables::CellMerge::Previous);
                cell->get_CellFormat()->set_VerticalMerge(currentPos.get_Y() == mergeRange.get_Y() ? Aspose::Words::Tables::CellMerge::First : Aspose::Words::Tables::CellMerge::Previous);
            }
        }
    }
}

namespace gtest_test
{

class ExTable : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExTable> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExTable>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExTable> ExTable::s_instance;

} // namespace gtest_test

void ExTable::CreateTable()
{
    //ExStart
    //ExFor:Table
    //ExFor:Row
    //ExFor:Cell
    //ExFor:Table.#ctor(DocumentBase)
    //ExSummary:Shows how to create a simple table.
    auto doc = MakeObject<Document>();

    // Tables are placed in the body of a document
    auto table = MakeObject<Table>(doc);
    doc->get_FirstSection()->get_Body()->AppendChild(table);

    // Tables contain rows, which contain cells,
    // which contain contents such as paragraphs, runs and even other tables
    // Calling table.EnsureMinimum will also make sure that a table has at least one row, cell and paragraph
    auto firstRow = MakeObject<Row>(doc);
    table->AppendChild(firstRow);

    auto firstCell = MakeObject<Cell>(doc);
    firstRow->AppendChild(firstCell);

    auto paragraph = MakeObject<Paragraph>(doc);
    firstCell->AppendChild(paragraph);

    auto run = MakeObject<Run>(doc, u"Hello world!");
    paragraph->AppendChild(run);

    doc->Save(ArtifactsDir + u"Table.CreateTable.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.CreateTable.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(1, table->get_Rows()->get_Count());
    ASSERT_EQ(1, table->get_FirstRow()->get_Cells()->get_Count());
    ASSERT_EQ(u"Hello world!\a\a", table->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExTable, CreateTable)
{
    s_instance->CreateTable();
}

} // namespace gtest_test

void ExTable::RowCellFormat()
{
    //ExStart
    //ExFor:Row.RowFormat
    //ExFor:RowFormat
    //ExFor:Cell.CellFormat
    //ExFor:CellFormat
    //ExFor:CellFormat.Shading
    //ExSummary:Shows how to modify the format of rows and cells.
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

    // The appearance of rows and individual cells can be edited using the respective formatting objects
    SharedPtr<Aspose::Words::Tables::RowFormat> rowFormat = table->get_FirstRow()->get_RowFormat();
    rowFormat->set_Height(25);
    rowFormat->get_Borders()->idx_get(Aspose::Words::BorderType::Bottom)->set_Color(System::Drawing::Color::get_Red());

    SharedPtr<Aspose::Words::Tables::CellFormat> cellFormat = table->get_LastRow()->get_FirstCell()->get_CellFormat();
    cellFormat->set_Width(100);
    cellFormat->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Orange());

    doc->Save(ArtifactsDir + u"Table.RowCellFormat.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.RowCellFormat.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(u"City\aCountry\a\aLondon\aU.K.\a\a", table->GetText().Trim());

    rowFormat = table->get_FirstRow()->get_RowFormat();

    ASPOSE_ASSERT_EQ(25.0, rowFormat->get_Height());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), rowFormat->get_Borders()->idx_get(Aspose::Words::BorderType::Bottom)->get_Color().ToArgb());

    cellFormat = table->get_LastRow()->get_FirstCell()->get_CellFormat();

    ASPOSE_ASSERT_EQ(110.8, cellFormat->get_Width());
    ASSERT_EQ(System::Drawing::Color::get_Orange().ToArgb(), cellFormat->get_Shading()->get_BackgroundPatternColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExTable, RowCellFormat)
{
    s_instance->RowCellFormat();
}

} // namespace gtest_test

void ExTable::DisplayContentOfTables()
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
    //ExSummary:Shows how to iterate through all tables in the document and display the content from each cell.
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

    // Here we get all tables from the Document node. You can do this for any other composite node
    // which can contain block level nodes. For example you can retrieve tables from header or from a cell
    // containing another table (nested tables)
    SharedPtr<TableCollection> tables = doc->get_FirstSection()->get_Body()->get_Tables();

    // We can make a new array to clone all of the tables in the collection
    ASSERT_EQ(2, tables->ToArray()->get_Length());

    // Iterate through all tables in the document
    for (int i = 0; i < tables->get_Count(); i++)
    {
        // Get the index of the table node as contained in the parent node of the table
        System::Console::WriteLine(String::Format(u"Start of Table {0}",i));

        SharedPtr<RowCollection> rows = tables->idx_get(i)->get_Rows();

        // RowCollections can be cloned into arrays
        ASPOSE_ASSERT_EQ(rows, rows->ToArray());
        ASPOSE_ASSERT_NS(rows, rows->ToArray());

        // Iterate through all rows in the table
        for (int j = 0; j < rows->get_Count(); j++)
        {
            System::Console::WriteLine(String::Format(u"\tStart of Row {0}",j));

            SharedPtr<CellCollection> cells = rows->idx_get(j)->get_Cells();

            // RowCollections can also be cloned into arrays
            ASPOSE_ASSERT_EQ(cells, cells->ToArray());
            ASPOSE_ASSERT_NS(cells, cells->ToArray());

            // Iterate through all cells in the row
            for (int k = 0; k < cells->get_Count(); k++)
            {
                // Get the plain text content of this cell
                String cellText = cells->idx_get(k)->ToString(Aspose::Words::SaveFormat::Text).Trim();
                // Print the content of the cell
                System::Console::WriteLine(String::Format(u"\t\tContents of Cell:{0} = \"{1}\"",k,cellText));
            }

            System::Console::WriteLine(String::Format(u"\tEnd of Row {0}",j));
        }

        System::Console::WriteLine(String::Format(u"End of Table {0}\n",i));
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, DisplayContentOfTables)
{
    s_instance->DisplayContentOfTables();
}

} // namespace gtest_test

void ExTable::CalculateDepthOfNestedTables()
{
    auto doc = MakeObject<Document>(MyDir + u"Nested tables.docx");
    SharedPtr<NodeCollection> tables = doc->GetChildNodes(Aspose::Words::NodeType::Table, true);
    ASSERT_EQ(5, tables->get_Count());
    //ExSkip

    for (int i = 0; i < tables->get_Count(); i++)
    {
        // First lets find if any cells in the table have tables themselves as children
        int count = GetChildTableCount(System::DynamicCast<Aspose::Words::Tables::Table>(tables->idx_get(i)));
        System::Console::WriteLine(u"Table #{0} has {1} tables directly within its cells", System::ObjectExt::Box<int>(i), System::ObjectExt::Box<int>(count));

        // Now let's try the other way around, lets try find if the table is nested inside another table and at what depth
        int tableDepth = GetNestedDepthOfTable(System::DynamicCast<Aspose::Words::Tables::Table>(tables->idx_get(i)));

        if (tableDepth > 0)
        {
            System::Console::WriteLine(u"Table #{0} is nested inside another table at depth of {1}", System::ObjectExt::Box<int>(i), System::ObjectExt::Box<int>(tableDepth));
        }
        else
        {
            System::Console::WriteLine(u"Table #{0} is a non nested table (is not a child of another table)", System::ObjectExt::Box<int>(i));
        }
    }
}

namespace gtest_test
{

TEST_F(ExTable, CalculateDepthOfNestedTables)
{
    s_instance->CalculateDepthOfNestedTables();
}

} // namespace gtest_test

void ExTable::ConvertTextBoxToTable()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a text box
    SharedPtr<Shape> textBox = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 300, 50);

    // Move the builder into the text box and write text
    builder->MoveTo(textBox->get_LastParagraph());
    builder->Write(u"Hello world!");

    // Convert all shape nodes which contain child nodes
    // We convert the collection to an array as static "snapshot" because the original textboxes will be removed after conversion which will
    // invalidate the enumerator
    for (auto shape : System::IterateOver(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->ToArray()->LINQ_OfType<SharedPtr<Shape> >()))
    {
        if (shape->get_HasChildNodes())
        {
            ConvertTextboxToTable(shape);
        }
    }

    doc->Save(ArtifactsDir + u"Table.ConvertTextBoxToTable.html");
}

namespace gtest_test
{

TEST_F(ExTable, ConvertTextBoxToTable)
{
    s_instance->ConvertTextBoxToTable();
}

} // namespace gtest_test

void ExTable::EnsureTableMinimum()
{
    //ExStart
    //ExFor:Table.EnsureMinimum
    //ExSummary:Shows how to ensure a table node is valid.
    auto doc = MakeObject<Document>();

    // Create a new table and add it to the document
    auto table = MakeObject<Table>(doc);
    doc->get_FirstSection()->get_Body()->AppendChild(table);

    // Currently, the table does not contain any rows, cells or nodes that can have content added to them
    ASSERT_EQ(0, table->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());

    // This method ensures that the table has one row, one cell and one paragraph; the minimal nodes required to begin editing
    table->EnsureMinimum();
    table->get_FirstRow()->get_FirstCell()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Hello world!"));
    //ExEnd

    ASSERT_EQ(4, table->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
}

namespace gtest_test
{

TEST_F(ExTable, EnsureTableMinimum)
{
    s_instance->EnsureTableMinimum();
}

} // namespace gtest_test

void ExTable::EnsureRowMinimum()
{
    //ExStart
    //ExFor:Row.EnsureMinimum
    //ExSummary:Shows how to ensure a row node is valid.
    auto doc = MakeObject<Document>();

    // Create a new table and add it to the document
    auto table = MakeObject<Table>(doc);
    doc->get_FirstSection()->get_Body()->AppendChild(table);

    // Create a new row and add it to the table
    auto row = MakeObject<Row>(doc);
    table->AppendChild(row);

    // Currently, the row does not contain any cells or nodes that can have content added to them
    ASSERT_EQ(0, row->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());

    // Ensure the row has at least one cell with one paragraph that we can edit
    row->EnsureMinimum();
    row->get_FirstCell()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Hello world!"));
    //ExEnd

    ASSERT_EQ(3, row->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
}

namespace gtest_test
{

TEST_F(ExTable, EnsureRowMinimum)
{
    s_instance->EnsureRowMinimum();
}

} // namespace gtest_test

void ExTable::EnsureCellMinimum()
{
    //ExStart
    //ExFor:Cell.EnsureMinimum
    //ExSummary:Shows how to ensure a cell node is valid.
    auto doc = MakeObject<Document>();

    // Create a new table and add it to the document
    auto table = MakeObject<Table>(doc);
    doc->get_FirstSection()->get_Body()->AppendChild(table);

    // Create a new row with a new cell and append it to the table
    auto row = MakeObject<Row>(doc);
    table->AppendChild(row);

    auto cell = MakeObject<Cell>(doc);
    row->AppendChild(cell);

    // Currently, the cell does not contain any cells or nodes that can have content added to them
    ASSERT_EQ(0, cell->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());

    // Ensure the cell has at least one paragraph that we can edit
    cell->EnsureMinimum();
    cell->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Hello world!"));
    //ExEnd

    ASSERT_EQ(2, cell->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
}

namespace gtest_test
{

TEST_F(ExTable, EnsureCellMinimum)
{
    s_instance->EnsureCellMinimum();
}

} // namespace gtest_test

void ExTable::SetOutlineBorders()
{
    //ExStart
    //ExFor:Table.Alignment
    //ExFor:TableAlignment
    //ExFor:Table.ClearBorders
    //ExFor:Table.ClearShading
    //ExFor:Table.SetBorder
    //ExFor:TextureIndex
    //ExFor:Table.SetShading
    //ExSummary:Shows how to apply a outline border to a table.
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    // Align the table to the center of the page
    table->set_Alignment(Aspose::Words::Tables::TableAlignment::Center);

    // Clear any existing borders and shading from the table
    table->ClearBorders();
    table->ClearShading();

    // Set a green border around the table but not inside
    table->SetBorder(Aspose::Words::BorderType::Left, Aspose::Words::LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
    table->SetBorder(Aspose::Words::BorderType::Right, Aspose::Words::LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
    table->SetBorder(Aspose::Words::BorderType::Top, Aspose::Words::LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
    table->SetBorder(Aspose::Words::BorderType::Bottom, Aspose::Words::LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);

    // Fill the cells with a light green solid color
    table->SetShading(Aspose::Words::TextureIndex::TextureSolid, System::Drawing::Color::get_LightGreen(), System::Drawing::Color::Empty);

    doc->Save(ArtifactsDir + u"Table.SetOutlineBorders.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.SetOutlineBorders.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(Aspose::Words::Tables::TableAlignment::Center, table->get_Alignment());

    SharedPtr<BorderCollection> borders = table->get_FirstRow()->get_RowFormat()->get_Borders();

    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), borders->get_Top()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), borders->get_Left()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), borders->get_Right()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), borders->get_Bottom()->get_Color().ToArgb());
    ASSERT_NE(System::Drawing::Color::get_Green().ToArgb(), borders->get_Horizontal()->get_Color().ToArgb());
    ASSERT_NE(System::Drawing::Color::get_Green().ToArgb(), borders->get_Vertical()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_LightGreen().ToArgb(), table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_ForegroundPatternColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExTable, SetOutlineBorders)
{
    s_instance->SetOutlineBorders();
}

} // namespace gtest_test

void ExTable::SetTableBorders()
{
    //ExStart
    //ExFor:Table.SetBorders
    //ExSummary:Shows how to build a table with all borders enabled (grid).
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    // Clear any existing borders from the table
    table->ClearBorders();

    // Set a green border around and inside the table
    table->SetBorders(Aspose::Words::LineStyle::Single, 1.5, System::Drawing::Color::get_Green());

    doc->Save(ArtifactsDir + u"Table.SetAllBorders.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.SetAllBorders.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Top()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Left()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Right()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Bottom()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Horizontal()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Vertical()->get_Color().ToArgb());
}

namespace gtest_test
{

TEST_F(ExTable, SetTableBorders)
{
    s_instance->SetTableBorders();
}

} // namespace gtest_test

void ExTable::RowFormat()
{
    //ExStart
    //ExFor:RowFormat
    //ExFor:Row.RowFormat
    //ExSummary:Shows how to modify formatting of a table row.
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    // Retrieve the first row in the table
    SharedPtr<Row> firstRow = table->get_FirstRow();

    // Modify some row level properties
    firstRow->get_RowFormat()->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::None);
    firstRow->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Auto);
    firstRow->get_RowFormat()->set_AllowBreakAcrossPages(true);

    doc->Save(ArtifactsDir + u"Table.RowFormat.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.RowFormat.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(Aspose::Words::LineStyle::None, table->get_FirstRow()->get_RowFormat()->get_Borders()->get_LineStyle());
    ASSERT_EQ(Aspose::Words::HeightRule::Auto, table->get_FirstRow()->get_RowFormat()->get_HeightRule());
    ASSERT_TRUE(table->get_FirstRow()->get_RowFormat()->get_AllowBreakAcrossPages());
}

namespace gtest_test
{

TEST_F(ExTable, RowFormat)
{
    s_instance->RowFormat();
}

} // namespace gtest_test

void ExTable::CellFormat()
{
    //ExStart
    //ExFor:CellFormat
    //ExFor:Cell.CellFormat
    //ExSummary:Shows how to modify formatting of a table cell.
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    // Retrieve the first cell in the table
    SharedPtr<Cell> firstCell = table->get_FirstRow()->get_FirstCell();

    // Modify some row level properties
    firstCell->get_CellFormat()->set_Width(30);
    // in points
    firstCell->get_CellFormat()->set_Orientation(Aspose::Words::TextOrientation::Downward);
    firstCell->get_CellFormat()->get_Shading()->set_ForegroundPatternColor(System::Drawing::Color::get_LightGreen());

    doc->Save(ArtifactsDir + u"Table.CellFormat.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.CellFormat.docx");

    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    ASPOSE_ASSERT_EQ(30, table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Width());
    ASSERT_EQ(Aspose::Words::TextOrientation::Downward, table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Orientation());
    ASSERT_EQ(System::Drawing::Color::get_LightGreen().ToArgb(), table->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Shading()->get_ForegroundPatternColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExTable, CellFormat)
{
    s_instance->CellFormat();
}

} // namespace gtest_test

void ExTable::GetDistance()
{
    //ExStart
    //ExFor:Table.DistanceBottom
    //ExFor:Table.DistanceLeft
    //ExFor:Table.DistanceRight
    //ExFor:Table.DistanceTop
    //ExSummary:Shows the minimum distance operations between table boundaries and text.
    auto doc = MakeObject<Document>(MyDir + u"Table wrapped by text.docx");

    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASPOSE_ASSERT_EQ(25.9, table->get_DistanceTop());
    ASPOSE_ASSERT_EQ(25.9, table->get_DistanceBottom());
    ASPOSE_ASSERT_EQ(17.3, table->get_DistanceLeft());
    ASPOSE_ASSERT_EQ(17.3, table->get_DistanceRight());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, GetDistance)
{
    s_instance->GetDistance();
}

} // namespace gtest_test

void ExTable::Borders()
{
    //ExStart
    //ExFor:Table.ClearBorders
    //ExSummary:Shows how to remove all borders from a table.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Create a table
    SharedPtr<Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Hello world!");
    builder->EndTable();

    // Set a color/thickness for the top border of the first row and verify the values
    SharedPtr<Border> topBorder = table->get_FirstRow()->get_RowFormat()->get_Borders()->idx_get(Aspose::Words::BorderType::Top);
    table->SetBorder(Aspose::Words::BorderType::Top, Aspose::Words::LineStyle::Double, 1.5, System::Drawing::Color::get_Red(), true);

    ASPOSE_ASSERT_EQ(1.5, topBorder->get_LineWidth());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), topBorder->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::LineStyle::Double, topBorder->get_LineStyle());

    // Clear the borders all cells in the table
    table->ClearBorders();
    doc->Save(ArtifactsDir + u"Table.ClearBorders.docx");

    // Upon re-opening the saved document, the new border attributes can be verified
    doc = MakeObject<Document>(ArtifactsDir + u"Table.ClearBorders.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    topBorder = table->get_FirstRow()->get_RowFormat()->get_Borders()->idx_get(Aspose::Words::BorderType::Top);

    ASPOSE_ASSERT_EQ(0.0, topBorder->get_LineWidth());
    ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), topBorder->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::LineStyle::None, topBorder->get_LineStyle());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, Borders)
{
    s_instance->Borders();
}

} // namespace gtest_test

void ExTable::ReplaceCellText()
{
    //ExStart
    //ExFor:Range.Replace(String, String, FindReplaceOptions)
    //ExSummary:Shows how to replace all instances of String of text in a table and cell.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Create a table and give it conditional styling on border colors based on the row being the first or last
    SharedPtr<Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Carrots");
    builder->InsertCell();
    builder->Write(u"30");
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"Potatoes");
    builder->InsertCell();
    builder->Write(u"50");
    builder->EndTable();

    auto options = MakeObject<FindReplaceOptions>();
    options->set_MatchCase(true);
    options->set_FindWholeWordsOnly(true);

    // Replace any instances of our String in the entire table
    table->get_Range()->Replace(u"Carrots", u"Eggs", options);
    // Replace any instances of our String in the last cell of the table only
    table->get_LastRow()->get_LastCell()->get_Range()->Replace(u"50", u"20", options);

    doc->Save(ArtifactsDir + u"Table.ReplaceCellText.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.ReplaceCellText.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(u"Eggs\a30\a\aPotatoes\a20\a\a", table->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExTable, ReplaceCellText)
{
    s_instance->ReplaceCellText();
}

} // namespace gtest_test

void ExTable::PrintTableRange()
{
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

    // Get the first table in the document
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    // The range text will include control characters such as "\a" for a cell
    // You can call ToString on the desired node to retrieve the plain text content

    // Print the plain text range of the table to the screen
    System::Console::WriteLine(u"Contents of the table: ");
    System::Console::WriteLine(table->get_Range()->get_Text());

    // Print the contents of the second row to the screen
    System::Console::WriteLine(u"\nContents of the row: ");
    System::Console::WriteLine(table->get_Rows()->idx_get(1)->get_Range()->get_Text());

    // Print the contents of the last cell in the table to the screen
    System::Console::WriteLine(u"\nContents of the cell: ");
    System::Console::WriteLine(table->get_LastRow()->get_LastCell()->get_Range()->get_Text());

    ASSERT_EQ(u"\aColumn 1\aColumn 2\aColumn 3\aColumn 4\a\a", table->get_Rows()->idx_get(1)->get_Range()->get_Text());
    ASSERT_EQ(u"Cell 12 contents\a", table->get_LastRow()->get_LastCell()->get_Range()->get_Text());
}

namespace gtest_test
{

TEST_F(ExTable, PrintTableRange)
{
    s_instance->PrintTableRange();
}

} // namespace gtest_test

void ExTable::CloneTable()
{
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

    // Retrieve the first table in the document
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    // Create a clone of the table
    auto tableClone = System::DynamicCast<Aspose::Words::Tables::Table>(table->Clone(true));

    // Insert the cloned table into the document after the original
    table->get_ParentNode()->InsertAfter(tableClone, table);

    // Insert an empty paragraph between the two tables or else they will be combined into one
    // upon save. This has to do with document validation
    table->get_ParentNode()->InsertAfter(MakeObject<Paragraph>(doc), table);

    doc->Save(ArtifactsDir + u"Table.CloneTable.doc");

    // Verify that the table was cloned and inserted properly
    ASSERT_EQ(3, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
    ASSERT_EQ(table->get_Range()->get_Text(), tableClone->get_Range()->get_Text());

    for (auto cell : System::IterateOver(tableClone->GetChildNodes(Aspose::Words::NodeType::Cell, true)->LINQ_OfType<SharedPtr<Cell> >()))
    {
        cell->RemoveAllChildren();
    }

    ASSERT_EQ(String::Empty, tableClone->ToString(Aspose::Words::SaveFormat::Text).Trim());
}

namespace gtest_test
{

TEST_F(ExTable, CloneTable)
{
    s_instance->CloneTable();
}

} // namespace gtest_test

void ExTable::DisableBreakAcrossPages()
{
    //ExStart
    //ExFor:RowFormat.AllowBreakAcrossPages
    //ExSummary:Shows how to disable rows breaking across pages for every row in a table.
    // Disable breaking across pages for all rows in the table
    auto doc = MakeObject<Document>(MyDir + u"Table spanning two pages.docx");

    // Retrieve the first table in the document
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    for (auto row : System::IterateOver(table->LINQ_OfType<SharedPtr<Row> >()))
    {
        row->get_RowFormat()->set_AllowBreakAcrossPages(false);
    }

    doc->Save(ArtifactsDir + u"Table.DisableBreakAcrossPages.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.DisableBreakAcrossPages.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_FALSE(table->get_FirstRow()->get_RowFormat()->get_AllowBreakAcrossPages());
    ASSERT_FALSE(table->get_LastRow()->get_RowFormat()->get_AllowBreakAcrossPages());
}

namespace gtest_test
{

TEST_F(ExTable, DisableBreakAcrossPages)
{
    s_instance->DisableBreakAcrossPages();
}

} // namespace gtest_test

void ExTable::AllowAutoFitOnTable(bool allowAutoFit)
{
    //ExStart
    //ExFor:Table.AllowAutoFit
    //ExSummary:Shows how to set a table to shrink or grow each cell to accommodate its contents.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<Table> table = builder->StartTable();
    builder->InsertCell();
    builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::FromPoints(100));
    builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

    builder->InsertCell();
    builder->get_CellFormat()->set_PreferredWidth(PreferredWidth::Auto());
    builder->Write(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    builder->EndRow();
    builder->EndTable();

    table->set_AllowAutoFit(allowAutoFit);

    doc->Save(ArtifactsDir + u"Table.AllowAutoFitOnTable.html");
    //ExEnd

    if (allowAutoFit)
    {
        TestUtil::FileContainsString(u"<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">", ArtifactsDir + u"Table.AllowAutoFitOnTable.html");
        TestUtil::FileContainsString(u"<td style=\"border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">", ArtifactsDir + u"Table.AllowAutoFitOnTable.html");
    }
    else
    {
        TestUtil::FileContainsString(u"<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">", ArtifactsDir + u"Table.AllowAutoFitOnTable.html");
        TestUtil::FileContainsString(u"<td style=\"width:7.2pt; border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">", ArtifactsDir + u"Table.AllowAutoFitOnTable.html");
    }
}

namespace gtest_test
{

using AllowAutoFitOnTable_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTable::AllowAutoFitOnTable)>::type;

struct AllowAutoFitOnTableVP : public ExTable, public ApiExamples::ExTable, public ::testing::WithParamInterface<AllowAutoFitOnTable_Args>
{
    static std::vector<AllowAutoFitOnTable_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(AllowAutoFitOnTableVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AllowAutoFitOnTable(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTable, AllowAutoFitOnTableVP, ::testing::ValuesIn(AllowAutoFitOnTableVP::TestCases()));

} // namespace gtest_test

void ExTable::KeepTableTogether()
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
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    // Enabling KeepWithNext for every paragraph in the table except for the last ones in the last row
    // will prevent the table from being split across pages
    for (auto cell : System::IterateOver(table->GetChildNodes(Aspose::Words::NodeType::Cell, true)->LINQ_OfType<SharedPtr<Cell> >()))
    {
        for (auto para : System::IterateOver(cell->get_Paragraphs()->LINQ_OfType<SharedPtr<Paragraph> >()))
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
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    for (auto para : System::IterateOver(table->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)->LINQ_OfType<SharedPtr<Paragraph> >()))
    {
        if (para->get_IsEndOfCell() && (System::DynamicCast<Aspose::Words::Tables::Cell>(para->get_ParentNode()))->get_ParentRow()->get_IsLastRow())
        {
            ASSERT_FALSE(para->get_ParagraphFormat()->get_KeepWithNext());
        }
        else
        {
            ASSERT_TRUE(para->get_ParagraphFormat()->get_KeepWithNext());
        }
    }
}

namespace gtest_test
{

TEST_F(ExTable, KeepTableTogether)
{
    s_instance->KeepTableTogether();
}

} // namespace gtest_test

void ExTable::FixDefaultTableWidthsInAw105()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Keep a reference to the table being built
    SharedPtr<Table> table = builder->StartTable();

    // Apply some formatting
    builder->get_CellFormat()->set_Width(100);
    builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Red());

    builder->InsertCell();
    // This will cause the table to be structured using column widths as in previous versions
    // instead of fitted to the page width like in the newer versions
    table->AutoFit(Aspose::Words::Tables::AutoFitBehavior::FixedColumnWidths);

    // Continue with building your table as usual...
}

namespace gtest_test
{

TEST_F(ExTable, FixDefaultTableWidthsInAw105)
{
    s_instance->FixDefaultTableWidthsInAw105();
}

} // namespace gtest_test

void ExTable::FixDefaultTableBordersIn105()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Keep a reference to the table being built
    SharedPtr<Table> table = builder->StartTable();

    builder->InsertCell();
    // Clear all borders to match the defaults used in previous versions
    table->ClearBorders();

    // Continue with building your table as usual...
}

namespace gtest_test
{

TEST_F(ExTable, FixDefaultTableBordersIn105)
{
    s_instance->FixDefaultTableBordersIn105();
}

} // namespace gtest_test

void ExTable::FixDefaultTableFormattingExceptionIn105()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Keep a reference to the table being built
    SharedPtr<Table> table = builder->StartTable();

    // We must first insert a new cell which in turn inserts a row into the table
    builder->InsertCell();
    // Once a row exists in our table we can apply table wide formatting
    table->set_AllowAutoFit(true);

    // Continue with building your table as usual...
}

namespace gtest_test
{

TEST_F(ExTable, FixDefaultTableFormattingExceptionIn105)
{
    s_instance->FixDefaultTableFormattingExceptionIn105();
}

} // namespace gtest_test

void ExTable::FixRowFormattingNotAppliedIn105()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->StartTable();

    // For the first row this will be set correctly
    builder->get_RowFormat()->set_HeadingFormat(true);

    builder->InsertCell();
    builder->Writeln(u"Text");
    builder->InsertCell();
    builder->Writeln(u"Text");

    // End the first row
    builder->EndRow();

    // Here we would normally define some other row formatting, such as disabling the
    // heading format. However at the moment this will be ignored and the value from the
    // first row reapplied to the row

    builder->InsertCell();

    // Instead make sure to specify the row formatting for the second row here
    builder->get_RowFormat()->set_HeadingFormat(false);

    // Continue with building your table as usual...
}

namespace gtest_test
{

TEST_F(ExTable, FixRowFormattingNotAppliedIn105)
{
    s_instance->FixRowFormattingNotAppliedIn105();
}

} // namespace gtest_test

void ExTable::GetIndexOfTableElements()
{
    //ExStart
    //ExFor:NodeCollection.IndexOf(Node)
    //ExSummary:Shows how to get the indexes of nodes in the collections that contain them.
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    SharedPtr<NodeCollection> allTables = doc->GetChildNodes(Aspose::Words::NodeType::Table, true);

    ASSERT_EQ(0, allTables->IndexOf(table));

    SharedPtr<Row> row = table->get_Rows()->idx_get(2);

    ASSERT_EQ(2, table->IndexOf(row));

    SharedPtr<Cell> cell = row->get_LastCell();

    ASSERT_EQ(4, row->IndexOf(cell));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, GetIndexOfTableElements)
{
    s_instance->GetIndexOfTableElements();
}

} // namespace gtest_test

void ExTable::GetPreferredWidthTypeAndValue()
{
    //ExStart
    //ExFor:PreferredWidthType
    //ExFor:PreferredWidth.Type
    //ExFor:PreferredWidth.Value
    //ExSummary:Shows how to verify the preferred width type of a table cell.
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

    // Find the first table in the document
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    SharedPtr<Cell> firstCell = table->get_FirstRow()->get_FirstCell();

    ASSERT_EQ(Aspose::Words::Tables::PreferredWidthType::Percent, firstCell->get_CellFormat()->get_PreferredWidth()->get_Type());
    ASPOSE_ASSERT_EQ(11.16, firstCell->get_CellFormat()->get_PreferredWidth()->get_Value());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, GetPreferredWidthTypeAndValue)
{
    s_instance->GetPreferredWidthTypeAndValue();
}

} // namespace gtest_test

void ExTable::AllowCellSpacing(bool allowCellSpacing)
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

    // Set the size of padding space between cells, and the switch that enables/negates this setting
    table->set_CellSpacing(3);
    table->set_AllowCellSpacing(allowCellSpacing);

    doc->Save(ArtifactsDir + u"Table.AllowCellSpacing.html");
    //ExEnd

    TestUtil::FileContainsString(allowCellSpacing ? u"<td style=\"border-style:solid; border-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">" : String(u"<td style=\"border-right-style:solid; border-right-width:0.75pt; border-bottom-style:solid; border-bottom-width:0.75pt; ") + u"padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">", ArtifactsDir + u"Table.AllowCellSpacing.html");
}

namespace gtest_test
{

using AllowCellSpacing_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTable::AllowCellSpacing)>::type;

struct AllowCellSpacingVP : public ExTable, public ApiExamples::ExTable, public ::testing::WithParamInterface<AllowCellSpacing_Args>
{
    static std::vector<AllowCellSpacing_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(AllowCellSpacingVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AllowCellSpacing(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTable, AllowCellSpacingVP, ::testing::ValuesIn(AllowCellSpacingVP::TestCases()));

} // namespace gtest_test

void ExTable::CreateNestedTable()
{
    auto doc = MakeObject<Document>();

    // Create the outer table with three rows and four columns
    SharedPtr<Table> outerTable = CreateTable(doc, 3, 4, u"Outer Table");
    // Add it to the document body
    doc->get_FirstSection()->get_Body()->AppendChild(outerTable);

    // Create another table with two rows and two columns
    SharedPtr<Table> innerTable = CreateTable(doc, 2, 2, u"Inner Table");
    // Add this table to the first cell of the outer table
    outerTable->get_FirstRow()->get_FirstCell()->AppendChild(innerTable);

    doc->Save(ArtifactsDir + u"Table.CreateNestedTable.docx");
    TestCreateNestedTable(MakeObject<Document>(ArtifactsDir + u"Table.CreateNestedTable.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExTable, CreateNestedTable)
{
    s_instance->CreateNestedTable();
}

} // namespace gtest_test

void ExTable::CheckCellsMerged()
{
    auto doc = MakeObject<Document>(MyDir + u"Table with merged cells.docx");

    // Retrieve the first table in the document
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    for (auto row : System::IterateOver(table->get_Rows()->LINQ_OfType<SharedPtr<Row> >()))
    {
        for (auto cell : System::IterateOver(row->get_Cells()->LINQ_OfType<SharedPtr<Cell> >()))
        {
            System::Console::WriteLine(PrintCellMergeType(cell));
        }
    }

    ASSERT_EQ(u"The cell at R1, C1 is vertically merged", PrintCellMergeType(table->get_FirstRow()->get_FirstCell()));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExTable, CheckCellsMerged)
{
    s_instance->CheckCellsMerged();
}

} // namespace gtest_test

void ExTable::MergeCellRange()
{
    // Open the document
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

    // Retrieve the first table in the body of the first section
    SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

    // We want to merge the range of cells found in between these two cells
    SharedPtr<Cell> cellStartRange = table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2);
    SharedPtr<Cell> cellEndRange = table->get_Rows()->idx_get(3)->get_Cells()->idx_get(3);

    // Merge all the cells between the two specified cells into one
    MergeCells(cellStartRange, cellEndRange);

    // Save the document
    doc->Save(ArtifactsDir + u"Table.MergeCellRange.doc");

    // Verify the cells were merged
    int mergedCellsCount = 0;
    for (auto node : System::IterateOver(table->GetChildNodes(Aspose::Words::NodeType::Cell, true)))
    {
        auto cell = System::DynamicCast<Aspose::Words::Tables::Cell>(node);
        if (cell->get_CellFormat()->get_HorizontalMerge() != Aspose::Words::Tables::CellMerge::None || cell->get_CellFormat()->get_VerticalMerge() != Aspose::Words::Tables::CellMerge::None)
        {
            mergedCellsCount++;
        }
    }

    ASSERT_EQ(4, mergedCellsCount);
    ASSERT_TRUE(table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2)->get_CellFormat()->get_HorizontalMerge() == Aspose::Words::Tables::CellMerge::First);
    ASSERT_TRUE(table->get_Rows()->idx_get(2)->get_Cells()->idx_get(2)->get_CellFormat()->get_VerticalMerge() == Aspose::Words::Tables::CellMerge::First);
    ASSERT_TRUE(table->get_Rows()->idx_get(3)->get_Cells()->idx_get(3)->get_CellFormat()->get_HorizontalMerge() == Aspose::Words::Tables::CellMerge::Previous);
    ASSERT_TRUE(table->get_Rows()->idx_get(3)->get_Cells()->idx_get(3)->get_CellFormat()->get_VerticalMerge() == Aspose::Words::Tables::CellMerge::Previous);
}

namespace gtest_test
{

TEST_F(ExTable, MergeCellRange)
{
    s_instance->MergeCellRange();
}

} // namespace gtest_test

void ExTable::CombineTables()
{
    //ExStart
    //ExFor:Cell.CellFormat
    //ExFor:CellFormat.Borders
    //ExFor:Table.Rows
    //ExFor:Table.FirstRow
    //ExFor:CellFormat.ClearFormatting
    //ExFor:CompositeNode.HasChildNodes
    //ExSummary:Shows how to combine the rows from two tables into one.
    // Load the document
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

    // Get the first and second table in the document
    // The rows from the second table will be appended to the end of the first table
    auto firstTable = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    auto secondTable = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));

    // Append all rows from the current table to the next
    // Due to the design of tables even tables with different cell count and widths can be joined into one table
    while (secondTable->get_HasChildNodes())
    {
        firstTable->get_Rows()->Add(secondTable->get_FirstRow());
    }

    // Remove the empty table container
    secondTable->Remove();

    doc->Save(ArtifactsDir + u"Table.CombineTables.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.CombineTables.docx");

    ASSERT_EQ(1, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
    ASSERT_EQ(9, doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_Rows()->get_Count());
    ASSERT_EQ(42, doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->GetChildNodes(Aspose::Words::NodeType::Cell, true)->get_Count());
}

namespace gtest_test
{

TEST_F(ExTable, CombineTables)
{
    s_instance->CombineTables();
}

} // namespace gtest_test

void ExTable::SplitTable()
{
    // Load the document
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");

    // Get the first table in the document
    auto firstTable = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    // We will split the table at the third row (inclusive)
    SharedPtr<Row> row = firstTable->get_Rows()->idx_get(2);

    // Create a new container for the split table
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(firstTable->Clone(false));

    // Insert the container after the original
    firstTable->get_ParentNode()->InsertAfter(table, firstTable);

    // Add a buffer paragraph to ensure the tables stay apart
    firstTable->get_ParentNode()->InsertAfter(MakeObject<Paragraph>(doc), firstTable);

    SharedPtr<Row> currentRow;

    do
    {
        currentRow = firstTable->get_LastRow();
        table->PrependChild(currentRow);
    } while (currentRow != row);

    doc->Save(ArtifactsDir + u"Table.SplitTable.docx");

    doc = MakeObject<Document>(ArtifactsDir + u"Table.SplitTable.docx");
    // Test we are adding the rows in the correct order and the
    // selected row was also moved
    ASPOSE_ASSERT_EQ(row, table->get_FirstRow());

    ASSERT_EQ(2, firstTable->get_Rows()->get_Count());
    ASSERT_EQ(3, table->get_Rows()->get_Count());
    ASSERT_EQ(3, doc->GetChildNodes(Aspose::Words::NodeType::Table, true)->get_Count());
}

namespace gtest_test
{

TEST_F(ExTable, SplitTable)
{
    s_instance->SplitTable();
}

} // namespace gtest_test

void ExTable::WrapText()
{
    //ExStart
    //ExFor:Table.TextWrapping
    //ExFor:TextWrapping
    //ExSummary:Shows how to work with table text wrapping.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a table and a paragraph of text after it
    SharedPtr<Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Cell 1");
    builder->InsertCell();
    builder->Write(u"Cell 2");
    builder->EndTable();
    table->set_PreferredWidth(PreferredWidth::FromPoints(300));

    builder->get_Font()->set_Size(16);
    builder->Writeln(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

    // Set the table to wrap text around it and push it down into the paragraph below be setting the position
    table->set_TextWrapping(Aspose::Words::Tables::TextWrapping::Around);
    table->set_AbsoluteHorizontalDistance(100);
    table->set_AbsoluteVerticalDistance(20);

    doc->Save(ArtifactsDir + u"Table.WrapText.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.WrapText.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(Aspose::Words::Tables::TextWrapping::Around, table->get_TextWrapping());
    ASPOSE_ASSERT_EQ(100.0, table->get_AbsoluteHorizontalDistance());
    ASPOSE_ASSERT_EQ(20.0, table->get_AbsoluteVerticalDistance());
}

namespace gtest_test
{

TEST_F(ExTable, WrapText)
{
    s_instance->WrapText();
}

} // namespace gtest_test

void ExTable::GetFloatingTableProperties()
{
    //ExStart
    //ExFor:Table.HorizontalAnchor
    //ExFor:Table.VerticalAnchor
    //ExFor:Table.AllowOverlap
    //ExFor:ShapeBase.AllowOverlap
    //ExSummary:Shows how get properties for floating tables
    auto doc = MakeObject<Document>(MyDir + u"Table wrapped by text.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(Aspose::Words::Tables::TextWrapping::Around, table->get_TextWrapping());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, table->get_HorizontalAnchor());
    ASSERT_EQ(Aspose::Words::Drawing::RelativeVerticalPosition::Paragraph, table->get_VerticalAnchor());
    ASPOSE_ASSERT_EQ(false, table->get_AllowOverlap());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, GetFloatingTableProperties)
{
    s_instance->GetFloatingTableProperties();
}

} // namespace gtest_test

void ExTable::ChangeFloatingTableProperties()
{
    //ExStart
    //ExFor:Table.RelativeHorizontalAlignment
    //ExFor:Table.RelativeVerticalAlignment
    //ExFor:Table.AbsoluteHorizontalDistance
    //ExFor:Table.AbsoluteVerticalDistance
    //ExSummary:Shows how set the location of floating tables.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a table
    SharedPtr<Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Table 1, cell 1");
    builder->EndTable();
    table->set_PreferredWidth(PreferredWidth::FromPoints(300));

    // We can set the table's location to a place on the page, such as the bottom right corner
    table->set_RelativeVerticalAlignment(Aspose::Words::Drawing::VerticalAlignment::Bottom);
    table->set_RelativeHorizontalAlignment(Aspose::Words::Drawing::HorizontalAlignment::Right);

    table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Table 2, cell 1");
    builder->EndTable();
    table->set_PreferredWidth(PreferredWidth::FromPoints(300));

    // We can also set a horizontal and vertical offset from the location in the paragraph where the table was inserted
    table->set_AbsoluteVerticalDistance(50);
    table->set_AbsoluteHorizontalDistance(100);

    doc->Save(ArtifactsDir + u"Table.ChangeFloatingTableProperties.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.ChangeFloatingTableProperties.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(Aspose::Words::Drawing::VerticalAlignment::Bottom, table->get_RelativeVerticalAlignment());
    ASSERT_EQ(Aspose::Words::Drawing::HorizontalAlignment::Right, table->get_RelativeHorizontalAlignment());

    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true));

    ASPOSE_ASSERT_EQ(50.0, table->get_AbsoluteVerticalDistance());
    ASPOSE_ASSERT_EQ(100.0, table->get_AbsoluteHorizontalDistance());
}

namespace gtest_test
{

TEST_F(ExTable, ChangeFloatingTableProperties)
{
    s_instance->ChangeFloatingTableProperties();
}

} // namespace gtest_test

void ExTable::TableStyleCreation()
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

    auto tableStyle = System::DynamicCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle1"));
    tableStyle->set_AllowBreakAcrossPages(true);
    tableStyle->set_Bidi(true);
    tableStyle->set_CellSpacing(5);
    tableStyle->set_BottomPadding(20);
    tableStyle->set_LeftPadding(5);
    tableStyle->set_RightPadding(10);
    tableStyle->set_TopPadding(20);
    tableStyle->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_AntiqueWhite());
    tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Blue());
    tableStyle->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::DotDash);

    table->set_Style(tableStyle);

    // Some Table attributes are linked to style variables
    ASSERT_TRUE(table->get_Bidi());
    ASPOSE_ASSERT_EQ(5.0, table->get_CellSpacing());
    ASSERT_EQ(u"MyTableStyle1", table->get_StyleName());

    doc->Save(ArtifactsDir + u"Table.TableStyleCreation.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.TableStyleCreation.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_TRUE(table->get_Bidi());
    ASPOSE_ASSERT_EQ(5.0, table->get_CellSpacing());
    ASSERT_EQ(u"MyTableStyle1", table->get_StyleName());
    ASPOSE_ASSERT_EQ(0.0, table->get_BottomPadding());
    ASPOSE_ASSERT_EQ(0.0, table->get_LeftPadding());
    ASPOSE_ASSERT_EQ(0.0, table->get_RightPadding());
    ASPOSE_ASSERT_EQ(0.0, table->get_TopPadding());

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Border> b)> _local_func_0 = [](SharedPtr<Border> b)
    {
        return b->get_Color().ToArgb() == System::Drawing::Color::get_Blue().ToArgb();
    };

    ASSERT_EQ(6, table->get_FirstRow()->get_RowFormat()->get_Borders()->LINQ_Count(static_cast<System::Func<SharedPtr<Border>, bool>>(_local_func_0)));

    tableStyle = System::DynamicCast<Aspose::Words::TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle1"));

    ASSERT_TRUE(tableStyle->get_AllowBreakAcrossPages());
    ASSERT_TRUE(tableStyle->get_Bidi());
    ASPOSE_ASSERT_EQ(5.0, tableStyle->get_CellSpacing());
    ASPOSE_ASSERT_EQ(20.0, tableStyle->get_BottomPadding());
    ASPOSE_ASSERT_EQ(5.0, tableStyle->get_LeftPadding());
    ASPOSE_ASSERT_EQ(10.0, tableStyle->get_RightPadding());
    ASPOSE_ASSERT_EQ(20.0, tableStyle->get_TopPadding());
    ASSERT_EQ(System::Drawing::Color::get_AntiqueWhite().ToArgb(), tableStyle->get_Shading()->get_BackgroundPatternColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), tableStyle->get_Borders()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::LineStyle::DotDash, tableStyle->get_Borders()->get_LineStyle());
}

namespace gtest_test
{

TEST_F(ExTable, TableStyleCreation)
{
    s_instance->TableStyleCreation();
}

} // namespace gtest_test

void ExTable::SetTableAlignment()
{
    //ExStart
    //ExFor:TableStyle.Alignment
    //ExFor:TableStyle.LeftIndent
    //ExSummary:Shows how to set table position.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // There are two ways of horizontally aligning a table using a custom table style
    // One way is to align it to a location on the page, such as the center
    auto tableStyle = System::DynamicCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle1"));
    tableStyle->set_Alignment(Aspose::Words::Tables::TableAlignment::Center);
    tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Blue());
    tableStyle->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::Single);

    // Insert a table and apply the style we created to it
    SharedPtr<Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Aligned to the center of the page");
    builder->EndTable();
    table->set_PreferredWidth(PreferredWidth::FromPoints(300));

    table->set_Style(tableStyle);

    // We can also set a specific left indent to the style, and apply it to the table
    tableStyle = System::DynamicCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle2"));
    tableStyle->set_LeftIndent(55);
    tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Green());
    tableStyle->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::Single);

    table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Aligned according to left indent");
    builder->EndTable();
    table->set_PreferredWidth(PreferredWidth::FromPoints(300));

    table->set_Style(tableStyle);

    doc->Save(ArtifactsDir + u"Table.TableStyleCreation.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.TableStyleCreation.docx");

    tableStyle = System::DynamicCast<Aspose::Words::TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle1"));

    ASSERT_EQ(Aspose::Words::Tables::TableAlignment::Center, tableStyle->get_Alignment());
    ASPOSE_ASSERT_EQ(tableStyle, (System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true)))->get_Style());

    tableStyle = System::DynamicCast<Aspose::Words::TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle2"));

    ASPOSE_ASSERT_EQ(55.0, tableStyle->get_LeftIndent());
    ASPOSE_ASSERT_EQ(tableStyle, (System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 1, true)))->get_Style());
}

namespace gtest_test
{

TEST_F(ExTable, SetTableAlignment)
{
    s_instance->SetTableAlignment();
}

} // namespace gtest_test

void ExTable::ConditionalStyles()
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

    // Create a table
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

    // Create a custom table style
    auto tableStyle = System::DynamicCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle1"));

    // Conditional styles are formatting changes that affect only some of the cells of the table based on a predicate,
    // such as the cells being in the last row
    // We can access these conditional styles by style type like this
    tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::FirstRow)->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_AliceBlue());

    // The same conditional style can be accessed by index
    tableStyle->get_ConditionalStyles()->idx_get(0)->get_Borders()->set_Color(System::Drawing::Color::get_Black());
    tableStyle->get_ConditionalStyles()->idx_get(0)->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::DotDash);
    ASSERT_EQ(Aspose::Words::ConditionalStyleType::FirstRow, tableStyle->get_ConditionalStyles()->idx_get(0)->get_Type());

    // It can also be found in the ConditionalStyles collection as an attribute
    tableStyle->get_ConditionalStyles()->get_FirstRow()->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);

    // Apply padding and text formatting to conditional styles
    tableStyle->get_ConditionalStyles()->get_LastRow()->set_BottomPadding(10);
    tableStyle->get_ConditionalStyles()->get_LastRow()->set_LeftPadding(10);
    tableStyle->get_ConditionalStyles()->get_LastRow()->set_RightPadding(10);
    tableStyle->get_ConditionalStyles()->get_LastRow()->set_TopPadding(10);
    tableStyle->get_ConditionalStyles()->get_LastColumn()->get_Font()->set_Bold(true);

    // List all possible style conditions
    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<ConditionalStyle>>> enumerator = tableStyle->get_ConditionalStyles()->GetEnumerator();
        while (enumerator->MoveNext())
        {
            SharedPtr<ConditionalStyle> currentStyle = enumerator->get_Current();
            if (currentStyle != nullptr)
            {
                System::Console::WriteLine(System::ObjectExt::Box<ConditionalStyleType>(currentStyle->get_Type()));
            }
        }
    }

    // Apply conditional style to the table
    table->set_Style(tableStyle);

    // Changes to the first row are enabled by the table's style options be default,
    // but need to be manually enabled for some other parts, such as the last column/row
    table->set_StyleOptions(table->get_StyleOptions() | Aspose::Words::Tables::TableStyleOptions::LastRow | Aspose::Words::Tables::TableStyleOptions::LastColumn);

    doc->Save(ArtifactsDir + u"Table.ConditionalStyles.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.ConditionalStyles.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    ASSERT_EQ(Aspose::Words::Tables::TableStyleOptions::Default | Aspose::Words::Tables::TableStyleOptions::LastRow | Aspose::Words::Tables::TableStyleOptions::LastColumn, table->get_StyleOptions());
    SharedPtr<ConditionalStyleCollection> conditionalStyles = (System::DynamicCast<Aspose::Words::TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle1")))->get_ConditionalStyles();

    ASSERT_EQ(Aspose::Words::ConditionalStyleType::FirstRow, conditionalStyles->idx_get(0)->get_Type());
    ASSERT_EQ(System::Drawing::Color::get_AliceBlue().ToArgb(), conditionalStyles->idx_get(0)->get_Shading()->get_BackgroundPatternColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_Black().ToArgb(), conditionalStyles->idx_get(0)->get_Borders()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::LineStyle::DotDash, conditionalStyles->idx_get(0)->get_Borders()->get_LineStyle());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, conditionalStyles->idx_get(0)->get_ParagraphFormat()->get_Alignment());

    ASSERT_EQ(Aspose::Words::ConditionalStyleType::LastRow, conditionalStyles->idx_get(2)->get_Type());
    ASPOSE_ASSERT_EQ(10.0, conditionalStyles->idx_get(2)->get_BottomPadding());
    ASPOSE_ASSERT_EQ(10.0, conditionalStyles->idx_get(2)->get_LeftPadding());
    ASPOSE_ASSERT_EQ(10.0, conditionalStyles->idx_get(2)->get_RightPadding());
    ASPOSE_ASSERT_EQ(10.0, conditionalStyles->idx_get(2)->get_TopPadding());

    ASSERT_EQ(Aspose::Words::ConditionalStyleType::LastColumn, conditionalStyles->idx_get(3)->get_Type());
    ASSERT_TRUE(conditionalStyles->idx_get(3)->get_Font()->get_Bold());
}

namespace gtest_test
{

TEST_F(ExTable, ConditionalStyles)
{
    s_instance->ConditionalStyles();
}

} // namespace gtest_test

void ExTable::ClearTableStyleFormatting()
{
    //ExStart
    //ExFor:ConditionalStyle.ClearFormatting
    //ExFor:ConditionalStyleCollection.ClearFormatting
    //ExSummary:Shows how to reset conditional table styles.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Create a table and give it conditional styling on border colors based on the row being the first or last
    builder->StartTable();
    builder->InsertCell();
    builder->Write(u"First row");
    builder->EndRow();
    builder->InsertCell();
    builder->Write(u"Last row");
    builder->EndTable();

    auto tableStyle = System::DynamicCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle1"));
    tableStyle->get_ConditionalStyles()->get_FirstRow()->get_Borders()->set_Color(System::Drawing::Color::get_Red());
    tableStyle->get_ConditionalStyles()->get_LastRow()->get_Borders()->set_Color(System::Drawing::Color::get_Blue());

    // Conditional styles can be cleared for specific parts of the table
    tableStyle->get_ConditionalStyles()->idx_get(0)->ClearFormatting();
    ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty, tableStyle->get_ConditionalStyles()->get_FirstRow()->get_Borders()->get_Color());

    // Also, they can be cleared for the entire table
    tableStyle->get_ConditionalStyles()->ClearFormatting();
    ASPOSE_ASSERT_EQ(System::Drawing::Color::Empty, tableStyle->get_ConditionalStyles()->get_LastRow()->get_Borders()->get_Color());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, ClearTableStyleFormatting)
{
    s_instance->ClearTableStyleFormatting();
}

} // namespace gtest_test

void ExTable::AlternatingRowStyles()
{
    //ExStart
    //ExFor:TableStyle.ColumnStripe
    //ExFor:TableStyle.RowStripe
    //ExSummary:Shows how to create conditional table styles that alternate between rows.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // The conditional style of a table can be configured to apply a different color to the row/column,
    // based on whether the row/column is even or odd, creating an alternating color pattern
    // We can also apply a number n to the row/column banding, meaning that the color alternates after every n rows/columns instead of one
    // Create a table where the columns will be banded by single columns and rows will banded in threes
    SharedPtr<Table> table = builder->StartTable();
    for (int i = 0; i < 15; i++)
    {
        for (int j = 0; j < 4; j++)
        {
            builder->InsertCell();
            builder->Writeln(String::Format(u"{0} column.",(j % 2 == 0 ? String(u"Even") : String(u"Odd"))));
            builder->Write(String::Format(u"Row banding {0}.",(i % 3 == 0 ? String(u"start") : String(u"continuation"))));
        }
        builder->EndRow();
    }
    builder->EndTable();

    // Set a line style for all the borders of the table
    auto tableStyle = System::DynamicCast<Aspose::Words::TableStyle>(doc->get_Styles()->Add(Aspose::Words::StyleType::Table, u"MyTableStyle1"));
    tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Black());
    tableStyle->get_Borders()->set_LineStyle(Aspose::Words::LineStyle::Double);

    // Set the two colors which will alternate over every 3 rows
    tableStyle->set_RowStripe(3);
    tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::OddRowBanding)->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightBlue());
    tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::EvenRowBanding)->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightCyan());

    // Set a color to apply to every even column, which will override any custom row coloring
    tableStyle->set_ColumnStripe(1);
    tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::EvenColumnBanding)->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_LightSalmon());

    // Apply the style to the table
    table->set_Style(tableStyle);

    // Row bands are automatically enabled, but column banding needs to be enabled manually like this
    // Row coloring will only be overridden if the column banding has been explicitly set a color
    table->set_StyleOptions(table->get_StyleOptions() | Aspose::Words::Tables::TableStyleOptions::ColumnBands);

    doc->Save(ArtifactsDir + u"Table.AlternatingRowStyles.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Table.AlternatingRowStyles.docx");
    table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    tableStyle = System::DynamicCast<Aspose::Words::TableStyle>(doc->get_Styles()->idx_get(u"MyTableStyle1"));

    ASPOSE_ASSERT_EQ(tableStyle, table->get_Style());
    ASSERT_EQ(table->get_StyleOptions() | Aspose::Words::Tables::TableStyleOptions::ColumnBands, table->get_StyleOptions());

    ASSERT_EQ(System::Drawing::Color::get_Black().ToArgb(), tableStyle->get_Borders()->get_Color().ToArgb());
    ASSERT_EQ(Aspose::Words::LineStyle::Double, tableStyle->get_Borders()->get_LineStyle());
    ASSERT_EQ(3, tableStyle->get_RowStripe());
    ASSERT_EQ(System::Drawing::Color::get_LightBlue().ToArgb(), tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::OddRowBanding)->get_Shading()->get_BackgroundPatternColor().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_LightCyan().ToArgb(), tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::EvenRowBanding)->get_Shading()->get_BackgroundPatternColor().ToArgb());
    ASSERT_EQ(1, tableStyle->get_ColumnStripe());
    ASSERT_EQ(System::Drawing::Color::get_LightSalmon().ToArgb(), tableStyle->get_ConditionalStyles()->idx_get(Aspose::Words::ConditionalStyleType::EvenColumnBanding)->get_Shading()->get_BackgroundPatternColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExTable, AlternatingRowStyles)
{
    s_instance->AlternatingRowStyles();
}

} // namespace gtest_test

void ExTable::ConvertToHorizontallyMergedCells()
{
    //ExStart
    //ExFor:Table.ConvertToHorizontallyMergedCells
    //ExSummary:Shows how to convert cells horizontally merged by width to cells merged by CellFormat.HorizontalMerge.
    auto doc = MakeObject<Document>(MyDir + u"Table with merged cells.docx");

    // MS Word does not write merge flags anymore, they define merged cells by its width
    // So AW by default define only 5 cells in a row and all of it didn't have horizontal merge flag
    SharedPtr<Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    SharedPtr<Row> row = table->get_Rows()->idx_get(0);
    ASSERT_EQ(5, row->get_Cells()->get_Count());

    // To resolve this inconvenience, we have added new public method to convert cells which are horizontally merged
    // by its width to the cell horizontally merged by flags. Thus now we have 7 cells and some of them have
    // horizontal merge value
    table->ConvertToHorizontallyMergedCells();
    row = table->get_Rows()->idx_get(0);
    ASSERT_EQ(7, row->get_Cells()->get_Count());

    ASSERT_EQ(Aspose::Words::Tables::CellMerge::None, row->get_Cells()->idx_get(0)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::First, row->get_Cells()->idx_get(1)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::Previous, row->get_Cells()->idx_get(2)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::None, row->get_Cells()->idx_get(3)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::First, row->get_Cells()->idx_get(4)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::Previous, row->get_Cells()->idx_get(5)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::None, row->get_Cells()->idx_get(6)->get_CellFormat()->get_HorizontalMerge());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExTable, ConvertToHorizontallyMergedCells)
{
    s_instance->ConvertToHorizontallyMergedCells();
}

} // namespace gtest_test

} // namespace ApiExamples