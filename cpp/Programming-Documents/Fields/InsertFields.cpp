#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>

using namespace Aspose::Words;

void InsertFields()
{
    std::cout << "InsertField example started." << std::endl;
    // ExStart:InsertField
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->InsertField(u"MERGEFIELD MyFieldName \\* MERGEFORMAT");
    System::String outputPath = outputDataDir + u"InsertField.docx";
    doc->Save(outputPath);
    // ExEnd:InsertField
    std::cout << "Inserted field in the document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertField example finished." << std::endl << std::endl;
}