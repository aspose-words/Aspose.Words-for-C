#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include "Model/Text/Range.h"
#include "Model/FindReplace/FindReplaceOptions.h"
#include "Model/FindReplace/FindReplaceDirection.h"


using namespace Aspose::Words;

void FindAndReplace()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_QuickStart();

    // Open the document.
    std::cout << (dataDir + u"ReplaceSimple.doc").ToUtf8String() << '\n';

    auto doc = System::MakeObject<Document>(dataDir + u"ReplaceSimple.doc");

    // Check the text of the document
    std::cout << "Original document text: " << doc->get_Range()->get_Text().ToUtf8String() << '\n';
    
    // Replace the text in the document.
    doc->get_Range()->Replace(u"_CustomerName_", u"James Bond", System::MakeObject<Replacing::FindReplaceOptions>(FindReplaceDirection::Forward));

    // Check the replacement was made.
    std::cout << "Document text after replace: " << doc->get_Range()->get_Text().ToUtf8String() << '\n';

    dataDir = dataDir + GetOutputFilePath(u"FindAndReplace.doc");
    // Save the modified document.
    doc->Save(dataDir);

    std::cout << "\nText found and replaced successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
