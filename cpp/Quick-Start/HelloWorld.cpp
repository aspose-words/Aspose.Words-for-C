#include "../examples.h"

#include "Model/Document/Document.h"
#include "Model/Document/DocumentBuilder.h"

using namespace Aspose::Words;

void HelloWorld()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_QuickStart();

    // Create a blank document.
    auto doc = System::MakeObject<Document>();

    // DocumentBuilder provides members to easily add content to a document.
    auto builder = System::MakeObject<DocumentBuilder>(doc);

    // Write a new paragraph in the document with the text "Hello World!"
    builder->Writeln(u"Hello World!");

    // Save the document in DOC format. The format to save as is inferred from the extension of the file name.
    // Aspose.Words supports saving any document in many more formats.
    dataDir = dataDir + GetOutputFilePath(u"HelloWorld.doc");
    doc->Save(dataDir);

    std::cout << "\nNew document created successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}