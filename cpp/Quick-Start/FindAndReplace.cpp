#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include "Model/Text/Range.h"
#include "Model/FindReplace/FindReplaceOptions.h"
#include "Model/FindReplace/FindReplaceDirection.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

void FindAndReplace()
{
    std::cout << "FindAndReplace example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_QuickStart();

    // Open the document.
    std::cout << (dataDir + u"ReplaceSimple.doc").ToUtf8String() << std::endl;

    auto doc = System::MakeObject<Document>(dataDir + u"ReplaceSimple.doc");

    // Check the text of the document
    std::cout << "Original document text: " << doc->get_Range()->get_Text().ToUtf8String() << std::endl;
    
    // Replace the text in the document.
    doc->get_Range()->Replace(u"_CustomerName_", u"James Bond", System::MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));

    // Check the replacement was made.
    std::cout << "Document text after replace: " << doc->get_Range()->get_Text().ToUtf8String() << std::endl;

    System::String outputPath = dataDir + GetOutputFilePath(u"FindAndReplace.doc");
    // Save the modified document.
    doc->Save(outputPath);
    std::cout << "Text found and replaced successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "FindAndReplace example finished." << std::endl << std::endl;
}
