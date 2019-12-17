#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>

using namespace Aspose::Words;

void ExtractTextOnly()
{
    std::cout << "ExtractTextOnly example started." << std::endl;
    // ExStart:ExtractTextOnly
    System::SharedPtr<Document> doc = System::MakeObject<Document>();

    // Enter a dummy field into the document.
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->InsertField(u"MERGEFIELD Field");

    // GetText will retrieve all field codes and special characters
    std::cout << "GetText() Result: " << doc->GetText().ToUtf8String() << std::endl;

    // ToString will export the node to the specified format. When converted to text it will not retrieve fields code 
    // Or special characters, but will still contain some natural formatting characters such as paragraph markers etc. 
    // This is the same as "viewing" the document as if it was opened in a text editor.
    std::cout << "ToString() Result: " << doc->ToString(SaveFormat::Text).ToUtf8String() << std::endl;
    // ExEnd:ExtractTextOnly
    std::cout << "ExtractTextOnly example finished." << std::endl << std::endl;
}