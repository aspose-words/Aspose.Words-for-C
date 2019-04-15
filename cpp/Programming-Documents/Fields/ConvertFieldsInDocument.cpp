#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Fields/FieldType.h>

#include "FieldsHelper.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void ConvertFieldsInDocument()
{
    std::cout << "ConvertFieldsInDocument example started." << std::endl;
    // ExStart:ConvertFieldsInDocument
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");

    // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to static text.
    ConvertFieldsToStaticText(doc, FieldType::FieldIf);

    System::String outputPath = dataDir + GetOutputFilePath(u"ConvertFieldsInDocument.doc");
    // Save the document with fields transformed to disk.
    doc->Save(outputPath);
    // ExEnd:ConvertFieldsInDocument
    std::cout << "Converted fields to static text in the document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ConvertFieldsInDocument example finished." << std::endl << std::endl;
}