#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>

#include "FieldsHelper.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void ConvertFieldsInDocument()
{
    std::cout << "ConvertFieldsInDocument example started." << std::endl;
    // ExStart:ConvertFieldsInDocument
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");

    // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to static text.
    ConvertFieldsToStaticText(doc, FieldType::FieldIf);

    System::String outputPath = outputDataDir + u"ConvertFieldsInDocument.doc";
    // Save the document with fields transformed to disk.
    doc->Save(outputPath);
    // ExEnd:ConvertFieldsInDocument
    std::cout << "Converted fields to static text in the document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ConvertFieldsInDocument example finished." << std::endl << std::endl;
}