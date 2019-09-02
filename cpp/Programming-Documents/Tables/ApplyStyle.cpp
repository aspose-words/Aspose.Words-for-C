#include "stdafx.h"
#include "examples.h"

#include <Model/Borders/BorderCollection.h>
#include <Model/Borders/LineStyle.h>
#include <Model/Borders/Shading.h>
#include <Model/Borders/TextureIndex.h>
#include "Model/Document/Document.h"
#include "Model/Document/DocumentBuilder.h"
#include <Model/Nodes/NodeType.h>
#include <Model/Styles/ConditionalStyleCollection.h>
#include <Model/Styles/Style.h>
#include <Model/Styles/StyleCollection.h>
#include <Model/Styles/StyleIdentifier.h>
#include <Model/Styles/StyleType.h>
#include <Model/Styles/TableStyle.h>
#include <Model/Tables/AutoFitBehavior.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/Table.h>
#include <Model/Tables/TableStyleOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace
{
    void BuildTableWithStyle(System::String const &outputDataDir)
    {
        // ExStart:BuildTableWithStyle
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        System::SharedPtr<Table> table = builder->StartTable();
        // We must insert at least one row first before setting any table formatting.
        builder->InsertCell();
        // Set the table style used based of the unique style identifier.
        // Note that not all table styles are available when saving as .doc format.
        table->set_StyleIdentifier(StyleIdentifier::MediumShading1Accent1);
        // Apply which features should be formatted by the style.
        table->set_StyleOptions(TableStyleOptions::FirstColumn | TableStyleOptions::RowBands | TableStyleOptions::FirstRow);
        table->AutoFit(AutoFitBehavior::AutoFitToContents);
        // Continue with building the table as normal.
        builder->Writeln(u"Item");
        builder->get_CellFormat()->set_RightPadding(40);
        builder->InsertCell();
        builder->Writeln(u"Quantity (kg)");
        builder->EndRow();
        builder->InsertCell();
        builder->Writeln(u"Apples");
        builder->InsertCell();
        builder->Writeln(u"20");
        builder->EndRow();
        builder->InsertCell();
        builder->Writeln(u"Bananas");
        builder->InsertCell();
        builder->Writeln(u"40");
        builder->EndRow();
        builder->InsertCell();
        builder->Writeln(u"Carrots");
        builder->InsertCell();
        builder->Writeln(u"50");
        builder->EndRow();
        System::String outputPath = outputDataDir + u"ApplyStyle.BuildTableWithStyle.docx";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:BuildTableWithStyle
        std::cout << "Table created successfully with table style." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ExpandFormattingOnCellsAndRowFromStyle(System::String const &inputDataDir)
    {
        // ExStart:ExpandFormattingOnCellsAndRowFromStyle
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Table.TableStyle.docx");
        // Get the first cell of the first table in the document.
        System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        System::SharedPtr<Cell> firstCell = table->get_FirstRow()->get_FirstCell();
        // First print the color of the cell shading. This should be empty as the current shading
        // Is stored in the table style.
        System::Drawing::Color cellShadingBefore = firstCell->get_CellFormat()->get_Shading()->get_BackgroundPatternColor();
        std::cout << "Cell shading before style expansion: " << cellShadingBefore.ToString().ToUtf8String() << std::endl;
        // Expand table style formatting to direct formatting.
        doc->ExpandTableStylesToDirectFormatting();
        // Now print the cell shading after expanding table styles. A blue background pattern color
        // Should have been applied from the table style.
        System::Drawing::Color cellShadingAfter = firstCell->get_CellFormat()->get_Shading()->get_BackgroundPatternColor();
        std::cout << "Cell shading after style expansion: " << cellShadingAfter.ToString().ToUtf8String() << std::endl;
        // ExEnd:ExpandFormattingOnCellsAndRowFromStyle
    }

    void CreateTableStyle(System::String const &outputDataDir)
    {
        // ExStart:CreateTableStyle
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Name");
        builder->InsertCell();
        builder->Write(u"Value");
        builder->EndRow();
        builder->InsertCell();
        builder->InsertCell();
        builder->EndTable();

        System::SharedPtr<TableStyle> tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->Add(StyleType::Table, u"MyTableStyle1"));
        tableStyle->get_Borders()->set_LineStyle(LineStyle::Double);
        tableStyle->get_Borders()->set_LineWidth(1);
        tableStyle->set_LeftPadding(18);
        tableStyle->set_RightPadding(18);
        tableStyle->set_TopPadding(12);
        tableStyle->set_BottomPadding(12);

        table->set_Style(tableStyle);

        System::String outputPath = outputDataDir + u"ApplyStyle.CreateTableStyle.docx";
        doc->Save(outputPath);
        // ExEnd:CreateTableStyle
    }

    void DefineConditionalFormatting(System::String const &outputDataDir)
    {
        // ExStart:DefineConditionalFormatting
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Name");
        builder->InsertCell();
        builder->Write(u"Value");
        builder->EndRow();
        builder->InsertCell();
        builder->InsertCell();
        builder->EndTable();

        System::SharedPtr<TableStyle> tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->Add(StyleType::Table, u"MyTableStyle1"));
        // Define background color to the first row of table.
        tableStyle->get_ConditionalStyles()->get_FirstRow()->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_GreenYellow());
        tableStyle->get_ConditionalStyles()->get_FirstRow()->get_Shading()->set_Texture(TextureIndex::TextureNone);

        table->set_Style(tableStyle);

        System::String outputPath = outputDataDir + u"ApplyStyle.DefineConditionalFormatting.docx";
        doc->Save(outputPath);
        // ExEnd:DefineConditionalFormatting
    }
}

void ApplyStyle()
{
    std::cout << "ApplyStyle example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithTables();
    System::String outputDataDir = GetOutputDataDir_WorkingWithTables();
    BuildTableWithStyle(outputDataDir);
    ExpandFormattingOnCellsAndRowFromStyle(inputDataDir);
    CreateTableStyle(outputDataDir);
    DefineConditionalFormatting(outputDataDir);
    std::cout << "ApplyStyle example finished." << std::endl << std::endl;
}