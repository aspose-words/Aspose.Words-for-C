#include "stdafx.h"
#include "examples.h"

#include <system/diagnostics/debug.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/AutoFitBehavior.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Model/Tables/PreferredWidthType.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

void AutoFitTableToWindow()
{
    std::cout << "AutoFitTableToWindow example started." << std::endl;
    // ExStart:AutoFitTableToPageWidth
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithTables();
    System::String outputDataDir = GetOutputDataDir_WorkingWithTables();
    // Open the document
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
    System::SharedPtr<Table> table = System::DynamicCast<Table>(doc->GetChild(NodeType::Table, 0, true));
    // Autofit the first table to the page width.
    table->AutoFit(AutoFitBehavior::AutoFitToWindow);
    System::String outputPath = outputDataDir + u"AutoFitTableToWindow.doc";
    // Save the document to disk.
    doc->Save(outputPath);
    System::SharedPtr<Table> firstTable = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    System::Diagnostics::Debug::Assert(firstTable->get_PreferredWidth()->get_Type() == PreferredWidthType::Percent, u"PreferredWidth type is not percent");
    System::Diagnostics::Debug::Assert(firstTable->get_PreferredWidth()->get_Value() == 100, u"PreferredWidth value is different than 100");
    // ExEnd:AutoFitTableToPageWidth
    std::cout << "Auto fit tables to window successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "AutoFitTableToWindow example finished." << std::endl << std::endl;
}