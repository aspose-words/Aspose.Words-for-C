#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceDirection.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

void FindAndReplace()
{
    std::cout << "FindAndReplace example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_QuickStart();
    System::String outputDataDir = GetOutputDataDir_QuickStart();

    // Open the document.
    System::String documentPath = inputDataDir + u"ReplaceSimple.doc";
    std::cout << documentPath.ToUtf8String() << std::endl;
    System::SharedPtr<Document> doc = System::MakeObject<Document>(documentPath);

    // Check the text of the document
    std::cout << "Original document text: " << doc->get_Range()->get_Text().ToUtf8String() << std::endl;
    
    // Replace the text in the document.
    doc->get_Range()->Replace(u"_CustomerName_", u"James Bond", System::MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));

    // Check the replacement was made.
    std::cout << "Document text after replace: " << doc->get_Range()->get_Text().ToUtf8String() << std::endl;

    System::String outputPath = outputDataDir + u"FindAndReplace.doc";
    // Save the modified document.
    doc->Save(outputPath);
    std::cout << "Text found and replaced successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "FindAndReplace example finished." << std::endl << std::endl;
}
