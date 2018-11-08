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
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void JoinContinuous()
{
    std::cout << "JoinContinuous example started." << std::endl;
    // ExStart:JoinContinuous
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");

    // Make the document appear straight after the destination documents content.
    srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(Aspose::Words::SectionStart::Continuous);

    // Append the source document using the original styles found in the source document.
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting);
    System::String outputPath = dataDir + GetOutputFilePath(u"JoinContinuous.doc");
    dstDoc->Save(outputPath);
    // ExEnd:JoinContinuous
    std::cout << "Document appended successfully with continous section start." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "JoinContinuous example finished." << std::endl << std::endl;
}
