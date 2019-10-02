#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Other/FieldUnknown.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void InsertFieldNone()
{
    std::cout << "InsertFieldNone example started." << std::endl;
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();
    // ExStart:InsertFieldNone
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    System::SharedPtr<FieldUnknown> field = System::DynamicCast<FieldUnknown>(builder->InsertField(FieldType::FieldNone, false));

    System::String outputPath = outputDataDir + u"InsertFieldNone.docx";
    doc->Save(outputPath);
    // ExEnd:InsertFieldNone
    std::cout << "Inserted field in the document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertFieldNone example finished." << std::endl << std::endl;
}