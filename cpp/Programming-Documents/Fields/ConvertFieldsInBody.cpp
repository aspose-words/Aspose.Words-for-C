#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Fields/FieldType.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/Section.h>

#include "FieldsHelper.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void ConvertFieldsInBody()
{
    std::cout << "ConvertFieldsInBody example started." << std::endl;
    // ExStart:ConvertFieldsInBody
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");

    // Pass the appropriate parameters to convert PAGE fields encountered to static text only in the body of the first section.
    ConvertFieldsToStaticText(doc->get_FirstSection()->get_Body(), FieldType::FieldPage);

    System::String outputPath = outputDataDir + u"ConvertFieldsInBody.doc";
    // Save the document with fields transformed to disk.
    doc->Save(outputPath);
    // ExEnd:ConvertFieldsInBody
    std::cout << "Converted fields to static text in the document body successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ConvertFieldsInBody example finished." << std::endl << std::endl;
}