#include "stdafx.h"
#include "examples.h"

#include <Model/Text/TextOrientation.h>
#include <Model/Text/HeightRule.h>
#include <Model/Tables/Table.h>
#include <Model/Tables/RowFormat.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/CellVerticalAlignment.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/AutoFitBehavior.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

void DocumentBuilderBuildTable()
{
    std::cout << "DocumentBuilderBuildTable example started." << std::endl;
    // ExStart:DocumentBuilderBuildTable
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    System::SharedPtr<Table> table = builder->StartTable();

    // Insert a cell
    builder->InsertCell();
    // Use fixed column widths.
    table->AutoFit(AutoFitBehavior::FixedColumnWidths);

    builder->get_CellFormat()->set_VerticalAlignment(CellVerticalAlignment::Center);
    builder->Write(u"This is row 1 cell 1");

    // Insert a cell
    builder->InsertCell();
    builder->Write(u"This is row 1 cell 2");

    builder->EndRow();

    // Insert a cell
    builder->InsertCell();

    // Apply new row formatting
    builder->get_RowFormat()->set_Height(100);
    builder->get_RowFormat()->set_HeightRule(HeightRule::Exactly);
    
    builder->get_CellFormat()->set_Orientation(TextOrientation::Upward);
    builder->Writeln(u"This is row 2 cell 1");

    // Insert a cell
    builder->InsertCell();
    builder->get_CellFormat()->set_Orientation(TextOrientation::Downward);
    builder->Writeln(u"This is row 2 cell 2");

    builder->EndRow();

    builder->EndTable();
    System::String outputPath = dataDir + GetOutputFilePath(u"DocumentBuilderBuildTable.doc");
    doc->Save(outputPath);
    // ExEnd:DocumentBuilderBuildTable
    std::cout << "Table build successfully using DocumentBuilder." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "DocumentBuilderBuildTable example finished." << std::endl << std::endl;
}
