#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/TableCollection.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/PreferredWidthType.h>
#include <Model/Tables/PreferredWidth.h>
#include <Model/Tables/AutoFitBehavior.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/Body.h>
#include <Model/Nodes/NodeType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

void AutoFitTableToFixedColumnWidths()
{
    std::cout << "AutoFitTableToFixedColumnWidths example started." << std::endl;
    // ExStart:AutoFitTableToFixedColumnWidths
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithTables();
    System::String outputDataDir = GetOutputDataDir_WorkingWithTables();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
    System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
    // Disable autofitting on this table.
    table->AutoFit(AutoFitBehavior::FixedColumnWidths);
    System::String outputPath = outputDataDir + u"AutoFitTableToFixedColumnWidths.doc";
    // Save the document to disk.
    doc->Save(outputPath);
    System::SharedPtr<Table> firstTable = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    System::Diagnostics::Debug::Assert(firstTable->get_PreferredWidth()->get_Type() == PreferredWidthType::Auto, u"PreferredWidth type is not auto");
    System::Diagnostics::Debug::Assert(firstTable->get_PreferredWidth()->get_Value() == 0, u"PreferredWidth value is not 0");
    System::Diagnostics::Debug::Assert(firstTable->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_Width() == 69.2, u"Cell width is not correct.");
    // ExEnd:AutoFitTableToFixedColumnWidths
    std::cout << "Auto fit tables to fixed column widths successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "AutoFitTableToFixedColumnWidths example finished." << std::endl << std::endl;
}