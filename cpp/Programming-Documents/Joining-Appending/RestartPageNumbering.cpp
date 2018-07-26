#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/console.h>
#include <Model/Sections/SectionStart.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words; 

void RestartPageNumbering()
{
    // ExStart:RestartPageNumbering
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();
    
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");
    
    // Set the appended document to appear on the next page.
    srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(Aspose::Words::SectionStart::NewPage);
    // Restart the page numbering for the document to be appended.
    srcDoc->get_FirstSection()->get_PageSetup()->set_RestartPageNumbering(true);
    
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting);
    dataDir = dataDir + GetOutputFilePath(u"RestartPageNumbering.doc");
    dstDoc->Save(dataDir);
    // ExEnd:RestartPageNumbering
    std::cout << "\nDocument appended successfully with restart page numbering.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}