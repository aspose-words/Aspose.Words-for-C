#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/console.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void UseDestinationStyles()
{
    // ExStart:UseDestinationStyles
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();
    
    // Load the documents to join.
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");
    
    // Append the source document using the styles of the destination document.
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::UseDestinationStyles);
    
    // Save the joined document to disk.
    dataDir = dataDir + GetOutputFilePath(u"UseDestinationStyles.doc");
    dstDoc->Save(dataDir);
    // ExEnd:UseDestinationStyles
    std::cout << "\nDocument appended successfully with use destination styles option.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

