#include "../../examples.h"

#include <system/string.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void SimpleAppendDocument()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();
    
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");
    
    // Append the source document to the destination document using no extra options.
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting);
    
    dataDir = dataDir + GetOutputFilePath(u"SimpleAppendDocument.doc");
    dstDoc->Save(dataDir);
    
    std::cout << "\nSimple document append.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
