#include "../../examples.h"

#include <system/shared_ptr.h>
#include <drawing/color.h>

#include "Model/Document/Document.h"
#include "Model/Document/DocumentBuilder.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/TableStyleOptions.h>
#include <Model/Tables/AutoFitBehavior.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Styles/StyleIdentifier.h>
#include <Model/Borders/Shading.h>
#include <Model/Nodes/NodeType.h>

using namespace Aspose::Words;

namespace
{
    void BuildTableWithStyle(System::String const &dataDir)
    {
        // ExStart:BuildTableWithStyle
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        auto table = builder->StartTable();
        // We must insert at least one row first before setting any table formatting.
        builder->InsertCell();
        // Set the table style used based of the unique style identifier.
        // Note that not all table styles are available when saving as .doc format.
        table->set_StyleIdentifier(Aspose::Words::StyleIdentifier::MediumShading1Accent1);
        // Apply which features should be formatted by the style.
        table->set_StyleOptions(Aspose::Words::Tables::TableStyleOptions::FirstColumn | Aspose::Words::Tables::TableStyleOptions::RowBands | Aspose::Words::Tables::TableStyleOptions::FirstRow);
        table->AutoFit(Aspose::Words::Tables::AutoFitBehavior::AutoFitToContents);
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
        System::String outputPath = dataDir + GetOutputFilePath(u"ApplyStyle.BuildTableWithStyle.doc");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:BuildTableWithStyle
        std::cout << "\nTable created successfully with table style.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void ExpandFormattingOnCellsAndRowFromStyle(System::String const &dataDir)
    {
        // ExStart:ExpandFormattingOnCellsAndRowFromStyle
        auto doc = System::MakeObject<Document>(dataDir + u"Table.TableStyle.doc");
        // Get the first cell of the first table in the document.
        auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
        auto firstCell = table->get_FirstRow()->get_FirstCell();
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
}

void ApplyStyle()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithTables();
    BuildTableWithStyle(dataDir);
    ExpandFormattingOnCellsAndRowFromStyle(dataDir);
}