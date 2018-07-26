#include "../../examples.h"

#include <system/shared_ptr.h>
#include <system/diagnostics/debug.h>

#include "Model/Document/Document.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/Row.h>
#include <Model/Tables/TableCollection.h>
#include <Model/Tables/AutoFitBehavior.h>
#include <Model/Tables/PreferredWidthType.h>
#include <Model/Tables/PreferredWidth.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/CellFormat.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/Body.h>
#include <Model/Nodes/NodeType.h>

using namespace Aspose::Words;

void AutoFitTableToContents()
{
    // ExStart:AutoFitTableToContents
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithTables();
    auto doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    // Auto fit the table to the cell contents
    table->AutoFit(Aspose::Words::Tables::AutoFitBehavior::AutoFitToContents);
    System::String outputPath = dataDir + GetOutputFilePath(u"AutoFitTableToContents.doc");
    // Save the document to disk.
    doc->Save(outputPath);
    System::Diagnostics::Debug::Assert(doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_PreferredWidth()->get_Type() == Aspose::Words::Tables::PreferredWidthType::Auto, u"PreferredWidth type is not auto");
    System::Diagnostics::Debug::Assert(doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_PreferredWidth()->get_Type() == Aspose::Words::Tables::PreferredWidthType::Auto, u"PrefferedWidth on cell is not auto");
    System::Diagnostics::Debug::Assert(doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_FirstRow()->get_FirstCell()->get_CellFormat()->get_PreferredWidth()->get_Value() == 0, u"PreferredWidth value is not 0");
    // ExEnd:AutoFitTableToContents
    std::cout << "\nAuto fit tables to contents successfully.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
}