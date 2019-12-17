#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Fields/FieldType.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/Section.h>
#include <Model/Text/Paragraph.h>

#include "FieldsHelper.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void ConvertFieldsInParagraph()
{
    std::cout << "ConvertFieldsInParagraph example started." << std::endl;
    // ExStart:ConvertFieldsInParagraph
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");

    // Pass the appropriate parameters to convert all IF fields to static text that are encountered only in the last 
    // Paragraph of the document.
    ConvertFieldsToStaticText(doc->get_FirstSection()->get_Body()->get_LastParagraph(), FieldType::FieldIf);

    System::String outputPath = outputDataDir + u"ConvertFieldsInParagraph.doc";
    // Save the document with fields transformed to disk.
    doc->Save(outputPath);
    // ExEnd:ConvertFieldsInParagraph
    std::cout << "Converted fields to static text in the paragraph successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ConvertFieldsInParagraph example finished." << std::endl << std::endl;
}