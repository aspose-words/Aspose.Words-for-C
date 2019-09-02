#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include "Model/Document/DocumentBuilder.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/TableAlignment.h>
#include <Model/Tables/RowFormat.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Tables/Cell.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Borders/BorderType.h>
#include <Model/Borders/BorderCollection.h>
#include <Model/Borders/Border.h>
#include <Model/Borders/LineStyle.h>
#include <Model/Borders/TextureIndex.h>
#include <Model/Borders/Shading.h>
#include <Model/Saving/OoxmlSaveOptions.h>
#include <Model/Saving/OoxmlCompliance.h>
#include <Model/Text/TextOrientation.h>
#include <Model/Text/HeightRule.h>
#include <Model/Settings/MsWordVersion.h>
#include <Model/Settings/CompatibilityOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
using namespace Aspose::Words::Tables;

namespace
{
    void ApplyOutlineBorder(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:ApplyOutlineBorder
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.EmptyTable.doc");
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // Align the table to the center of the page.
        table->set_Alignment(TableAlignment::Center);
        // Clear any existing borders from the table.
        table->ClearBorders();
        // Set a green border around the table but not inside.
        table->SetBorder(BorderType::Left, LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
        table->SetBorder(BorderType::Right, LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
        table->SetBorder(BorderType::Top, LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
        table->SetBorder(BorderType::Bottom, LineStyle::Single, 1.5, System::Drawing::Color::get_Green(), true);
        // Fill the cells with a light green solid color.
        table->SetShading(TextureIndex::TextureSolid, System::Drawing::Color::get_LightGreen(), System::Drawing::Color::Empty);
        System::String outputPath = outputDataDir + u"ApplyFormatting.ApplyOutlineBorder.doc";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:ApplyOutlineBorder
        std::cout << "Outline border applied successfully to a table." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void BuildTableWithBordersEnabled(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:BuildTableWithBordersEnabled
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.EmptyTable.doc");
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // Clear any existing borders from the table.
        table->ClearBorders();
        // Set a green border around and inside the table.
        table->SetBorders(LineStyle::Single, 1.5, System::Drawing::Color::get_Green());
        System::String outputPath = outputDataDir + u"ApplyFormatting.BuildTableWithBordersEnabled.doc";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:BuildTableWithBordersEnabled
        std::cout << "Table build successfully with all borders enabled." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;

    }

    void ModifyRowFormatting(System::String const &inputDataDir)
    {
        // ExStart:ModifyRowFormatting
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.Document.doc");
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // Retrieve the first row in the table.
        System::SharedPtr<Row> firstRow = table->get_FirstRow();
        // Modify some row level properties.
        firstRow->get_RowFormat()->get_Borders()->set_LineStyle(LineStyle::None);
        firstRow->get_RowFormat()->set_HeightRule(HeightRule::Auto);
        firstRow->get_RowFormat()->set_AllowBreakAcrossPages(true);
        // ExEnd:ModifyRowFormatting
        std::cout << "Some row level properties modified successfully." << std::endl;
    }

    void ApplyRowFormatting(System::String const &outputDataDir)
    {
        // ExStart:ApplyRowFormatting
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        System::SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        // Set the row formatting
        System::SharedPtr<RowFormat> rowFormat = builder->get_RowFormat();
        rowFormat->set_Height(100);
        rowFormat->set_HeightRule(HeightRule::Exactly);
        // These formatting properties are set on the table and are applied to all rows in the table.
        table->set_LeftPadding(30);
        table->set_RightPadding(30);
        table->set_TopPadding(30);
        table->set_BottomPadding(30);
        builder->Writeln(u"I'm a wonderful formatted row.");
        builder->EndRow();
        builder->EndTable();
        System::String outputPath = outputDataDir + u"ApplyFormatting.ApplyRowFormatting.doc";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:ApplyRowFormatting
        std::cout << "Row formatting applied successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ModifyCellFormatting(System::String const &inputDataDir)
    {
        // ExStart:ModifyCellFormatting
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.Document.doc");
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // Retrieve the first cell in the table.
        System::SharedPtr<Cell> firstCell = table->get_FirstRow()->get_FirstCell();
        // Modify some cell level properties.
        firstCell->get_CellFormat()->set_Width(30);
        // In points
        firstCell->get_CellFormat()->set_Orientation(TextOrientation::Downward);
        firstCell->get_CellFormat()->get_Shading()->set_ForegroundPatternColor(System::Drawing::Color::get_LightGreen());
        // ExEnd:ModifyCellFormatting
        std::cout << "Some cell level properties modified successfully." << std::endl;
    }

    void FormatTableAndCellWithDifferentBorders(System::String const &outputDataDir)
    {
        // ExStart:FormatTableAndCellWithDifferentBorders
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        System::SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        // Set the borders for the entire table.
        table->SetBorders(LineStyle::Single, 2.0, System::Drawing::Color::get_Black());
        // Set the cell shading for this cell.
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Red());
        builder->Writeln(u"Cell #1");
        builder->InsertCell();
        // Specify a different cell shading for the second cell.
        builder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_Green());
        builder->Writeln(u"Cell #2");
        // End this row.
        builder->EndRow();
        // Clear the cell formatting from previous operations.
        builder->get_CellFormat()->ClearFormatting();
        // Create the second row.
        builder->InsertCell();
        // Create larger borders for the first cell of this row. This will be different.
        // Compared to the borders set for the table.
        builder->get_CellFormat()->get_Borders()->get_Left()->set_LineWidth(4.0);
        builder->get_CellFormat()->get_Borders()->get_Right()->set_LineWidth(4.0);
        builder->get_CellFormat()->get_Borders()->get_Top()->set_LineWidth(4.0);
        builder->get_CellFormat()->get_Borders()->get_Bottom()->set_LineWidth(4.0);
        builder->Writeln(u"Cell #3");
        builder->InsertCell();
        // Clear the cell formatting from the previous cell.
        builder->get_CellFormat()->ClearFormatting();
        builder->Writeln(u"Cell #4");
        // Save finished document.
        System::String outputPath = outputDataDir + u"ApplyFormatting.FormatTableAndCellWithDifferentBorders.doc";
        doc->Save(outputPath);
        // ExEnd:FormatTableAndCellWithDifferentBorders
        std::cout << "Format table and cell with different borders and shadings successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetCellPadding(System::String const &outputDataDir)
    {
        // ExStart:SetCellPadding
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        builder->StartTable();
        builder->InsertCell();
        //Sets the amount of space (in points) to add to the left/top/right/bottom of the contents of cell.
        builder->get_CellFormat()->SetPaddings(30, 50, 30, 50);
        builder->Writeln(u"I'm a wonderful formatted cell.");
        builder->EndRow();
        builder->EndTable();
        System::String outputPath = outputDataDir + u"ApplyFormatting.SetCellPadding.doc";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetCellPadding
        std::cout << "Cell padding is set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void GetDistancebetweenTableSurroundingText(System::String const &inputDataDir)
    {
        // ExStart:GetDistancebetweenTableSurroundingText
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.EmptyTable.doc");
        std::cout << "Get distance between table left, right, bottom, top and the surrounding text." << std::endl;
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        std::cout << table->get_DistanceTop() << std::endl;
        std::cout << table->get_DistanceBottom() << std::endl;
        std::cout << table->get_DistanceRight() << std::endl;
        std::cout << table->get_DistanceLeft() << std::endl;
        // ExEnd:GetDistancebetweenTableSurroundingText
    }

    void SetTableTitleAndDescription(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetTableTitleandDescription
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.Document.doc");
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        table->set_Title(u"Test title");
        table->set_Description(u"Test description");
        System::SharedPtr<OoxmlSaveOptions> options = System::MakeObject<OoxmlSaveOptions>();
        options->set_Compliance(OoxmlCompliance::Iso29500_2008_Strict);
        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2016);
        System::String outputPath = outputDataDir + u"ApplyFormatting.SetTableTitleandDescription.docx";
        // Save the document to disk.
        doc->Save(outputPath, options);
        // ExEnd:SetTableTitleandDescription
        std::cout << "Table's title and description is set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void AllowCellSpacing(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:AllowCellSpacing
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.Document.doc");
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        table->set_AllowCellSpacing(true);
        table->set_CellSpacing(2);
        System::String outputPath = outputDataDir + u"ApplyFormatting.AllowCellSpacing.docx";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:AllowCellSpacing
        std::cout << "Allow spacing between cells is set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void ApplyFormatting()
{
    std::cout << "ApplyFormatting example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithTables();
    System::String outputDataDir = GetOutputDataDir_WorkingWithTables();
    ApplyOutlineBorder(inputDataDir, outputDataDir);
    BuildTableWithBordersEnabled(inputDataDir, outputDataDir);
    ModifyRowFormatting(inputDataDir);
    ApplyRowFormatting(outputDataDir);
    ModifyCellFormatting(inputDataDir);
    FormatTableAndCellWithDifferentBorders(outputDataDir);
    SetCellPadding(outputDataDir);
    //Get DistanceLeft, DistanceRight, DistanceTop, and DistanceBottom properties
    GetDistancebetweenTableSurroundingText(inputDataDir);
    SetTableTitleAndDescription(inputDataDir, outputDataDir);
    AllowCellSpacing(inputDataDir, outputDataDir);
    std::cout << "ApplyFormatting example finished." << std::endl << std::endl;
}