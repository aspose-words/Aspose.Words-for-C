#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include "Model/Document/DocumentBuilder.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Tables/PreferredWidth.h>
#include <Model/Tables/PreferredWidthType.h>
#include <Model/Borders/Shading.h>
#include <Model/Nodes/NodeType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

namespace
{
    void AutoFitToPageWidth(System::String const &dataDir)
    {
        // ExStart:AutoFitToPageWidth
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        // Insert a table with a width that takes up half the page width.
        auto table = builder->StartTable();
        // Insert a few cells
        builder->InsertCell();
        table->set_PreferredWidth(PreferredWidth::FromPercent(50));
        builder->Writeln(u"Cell #1");
        builder->InsertCell();
        builder->Writeln(u"Cell #2");
        builder->InsertCell();
        builder->Writeln(u"Cell #3");
        System::String outputPath = dataDir + GetOutputFilePath(u"SpecifyHeightAndWidth.AutoFitToPageWidth.doc");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:AutoFitToPageWidth
        std::cout << "Table autofit successfully to 50% of the page width." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetPreferredWidthSettings(System::String const &dataDir)
    {
        // ExStart:SetPreferredWidthSettings
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        // Insert a table row made up of three cells which have different preferred widths.
        auto table = builder->StartTable();
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
        System::String outputPath = dataDir + GetOutputFilePath(u"SpecifyHeightAndWidth.SetPreferredWidthSettings.doc");
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetPreferredWidthSettings
        std::cout << "Different preferred width settings set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void RetrievePreferredWidthType(System::String const &dataDir)
    {
        // ExStart:RetrievePreferredWidthType
        auto doc = System::MakeObject<Document>(dataDir + u"Table.SimpleTable.doc");
        // Retrieve the first table in the document.
        auto table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
        // ExStart:AllowAutoFit
        table->set_AllowAutoFit(true);
        // ExEnd:AllowAutoFit
        auto firstCell = table->get_FirstRow()->get_FirstCell();
        PreferredWidthType type = firstCell->get_CellFormat()->get_PreferredWidth()->get_Type();
        double value = firstCell->get_CellFormat()->get_PreferredWidth()->get_Value();
        // ExEnd:RetrievePreferredWidthType
        std::cout << "Table preferred width type value is " << value << std::endl;
    }
}

void SpecifyHeightAndWidth()
{
    std::cout << "SpecifyHeightAndWidth example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithTables();
    AutoFitToPageWidth(dataDir);
    SetPreferredWidthSettings(dataDir);
    RetrievePreferredWidthType(dataDir);
    std::cout << "SpecifyHeightAndWidth example finished." << std::endl << std::endl;
}