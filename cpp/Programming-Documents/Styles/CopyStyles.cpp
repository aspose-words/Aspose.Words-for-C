#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>

using namespace Aspose::Words;

void CopyStyles()
{
    std::cout << "CopyStyles example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithStyles();
    System::String outputDataDir = GetOutputDataDir_WorkingWithStyles();

    // ExStart:CopyStylesFromDocument
    System::String fileName = inputDataDir + u"template.docx";
    System::SharedPtr<Document> doc = System::MakeObject<Document>(fileName);

    // Open the document.
    System::SharedPtr<Document> target = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
    target->CopyStylesFromTemplate(doc);
    System::String outputPath = outputDataDir + u"CopyStyles.doc";
    doc->Save(outputPath);
    // ExEnd:CopyStylesFromDocument 
    std::cout << "Styles are copied from document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CopyStyles example finished." << std::endl << std::endl;
}