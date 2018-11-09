#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void UpdateDocFields()
{
    std::cout << "UpdateDocFields example started." << std::endl;
    // ExStart:UpdateDocFields
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithFields();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Rendering.doc");

    // This updates all fields in the document.
    doc->UpdateFields();
    System::String outputPath = dataDir + GetOutputFilePath(u"UpdateDocFields.docx");
    doc->Save(outputPath);
    // ExEnd:UpdateDocFields
    std::cout << "Document fields updated successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "UpdateDocFields example finished." << std::endl << std::endl;
}