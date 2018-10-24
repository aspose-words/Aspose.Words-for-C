#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/console.h>
#include <Model/Sections/SectionStart.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Sections/Orientation.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void DifferentPageSetup()
{
    std::cout << "DifferentPageSetup example started." << std::endl;
    // ExStart:DifferentPageSetup
    // The path to the documents directory.
    System::String dataDir = ::GetDataDir_JoiningAndAppending();
    
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.SourcePageSetup.doc");
    
    // Set the source document to continue straight after the end of the destination document.
    // If some page setup settings are different then this may not work and the source document will appear 
    // On a new page.
    srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(Aspose::Words::SectionStart::Continuous);
    
    // To ensure this does not happen when the source document has different page setup settings make sure the
    // Settings are identical between the last section of the destination document.
    // If there are further continuous sections that follow on in the source document then this will need to be 
    // Repeated for those sections as well.
    srcDoc->get_FirstSection()->get_PageSetup()->set_PageWidth(dstDoc->get_LastSection()->get_PageSetup()->get_PageWidth());
    srcDoc->get_FirstSection()->get_PageSetup()->set_PageHeight(dstDoc->get_LastSection()->get_PageSetup()->get_PageHeight());
    srcDoc->get_FirstSection()->get_PageSetup()->set_Orientation(dstDoc->get_LastSection()->get_PageSetup()->get_Orientation());
    
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting);
    System::String outputPath = dataDir + GetOutputFilePath(u"DifferentPageSetup.doc");
    dstDoc->Save(outputPath);
    // ExEnd:DifferentPageSetup
    std::cout << "Document appended successfully with different page setup." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "DifferentPageSetup example finished." << std::endl << std::endl;
}
