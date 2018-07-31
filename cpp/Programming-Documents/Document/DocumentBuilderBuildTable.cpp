#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Text/TextOrientation.h>
#include <Model/Text/HeightRule.h>
#include <Model/Tables/Table.h>
#include <Model/Tables/RowFormat.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/CellVerticalAlignment.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/AutoFitBehavior.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

void DocumentBuilderBuildTable()
{
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
    table->AutoFit(Aspose::Words::Tables::AutoFitBehavior::FixedColumnWidths);
    
    builder->get_CellFormat()->set_VerticalAlignment(Aspose::Words::Tables::CellVerticalAlignment::Center);
    builder->Write(u"This is row 1 cell 1");
    
    // Insert a cell
    builder->InsertCell();
    builder->Write(u"This is row 1 cell 2");
    
    builder->EndRow();
    
    // Insert a cell
    builder->InsertCell();
    
    // Apply new row formatting
    builder->get_RowFormat()->set_Height(100);
    builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Exactly);
    
    builder->get_CellFormat()->set_Orientation(Aspose::Words::TextOrientation::Upward);
    builder->Writeln(u"This is row 2 cell 1");
    
    // Insert a cell
    builder->InsertCell();
    builder->get_CellFormat()->set_Orientation(Aspose::Words::TextOrientation::Downward);
    builder->Writeln(u"This is row 2 cell 2");
    
    builder->EndRow();
    
    builder->EndTable();
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderBuildTable.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderBuildTable
    std::cout << "\nTable build successfully using DocumentBuilder.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
