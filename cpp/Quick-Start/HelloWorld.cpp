#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include "Model/Document/DocumentBuilder.h"

using namespace Aspose::Words;

void HelloWorld()
{
    std::cout << "HelloWorld example started." << std::endl;
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_QuickStart();

    // Create a blank document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();

    // DocumentBuilder provides members to easily add content to a document.
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Write a new paragraph in the document with the text "Hello World!"
    builder->Writeln(u"Hello World!");

    // Save the document in DOC format. The format to save as is inferred from the extension of the file name.
    // Aspose.Words supports saving any document in many more formats.
    System::String outputPath = outputDataDir + u"HelloWorld.docx";
    doc->Save(outputPath);
    std::cout << "New document created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "HelloWorld example finished." << std::endl << std::endl;
}