#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/console.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/HeaderFooterCollection.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void UnlinkHeadersFooters()
{
    // ExStart:UnlinkHeadersFooters
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();
    
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");
    
    // Unlink the headers and footers in the source document to stop this from continuing the headers and footers
    // From the destination document.
    srcDoc->get_FirstSection()->get_HeadersFooters()->LinkToPrevious(false);
    
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting);
    dataDir = dataDir + GetOutputFilePath(u"UnlinkHeadersFooters.doc");
    dstDoc->Save(dataDir);
    // ExEnd:UnlinkHeadersFooters
    std::cout << "\nDocument appended successfully with unlinked header footers.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
