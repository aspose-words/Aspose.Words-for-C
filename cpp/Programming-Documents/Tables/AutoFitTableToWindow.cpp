#include "../../examples.h"

#include <system/shared_ptr.h>
#include <system/diagnostics/debug.h>

#include "Model/Document/Document.h"
#include <Model/Tables/Table.h>
#include <Model/Tables/TableCollection.h>
#include <Model/Tables/AutoFitBehavior.h>
#include <Model/Tables/PreferredWidth.h>
#include <Model/Tables/PreferredWidthType.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/Body.h>
#include <Model/Nodes/NodeType.h>

using namespace Aspose::Words;

void AutoFitTableToWindow()
{
    // ExStart:AutoFitTableToPageWidth
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithTables();
    // Open the document
    auto doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    // Autofit the first table to the page width.
    table->AutoFit(Aspose::Words::Tables::AutoFitBehavior::AutoFitToWindow);
    System::String outputPath = dataDir + GetOutputFilePath(u"AutoFitTableToWindow.doc");
    // Save the document to disk.
    doc->Save(outputPath);
    System::Diagnostics::Debug::Assert(doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_PreferredWidth()->get_Type() == Aspose::Words::Tables::PreferredWidthType::Percent, u"PreferredWidth type is not percent");
    System::Diagnostics::Debug::Assert(doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->get_PreferredWidth()->get_Value() == 100, u"PreferredWidth value is different than 100");
    // ExEnd:AutoFitTableToPageWidth
    std::cout << "\nAuto fit tables to window successfully.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
}