#include "stdafx.h"
#include "examples.h"

#include "system/shared_ptr.h"

#include "Model/Document/Document.h"
#include "Model/Importing/ImportFormatMode.h"

using namespace Aspose::Words;

void AppendDocuments()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_QuickStart();
    // Load the destination and source documents from disk.
    auto dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    auto srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");

    // Append the source document to the destination document while keeping the original formatting of the source document.
    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

    dataDir = dataDir + GetOutputFilePath(u"AppendDocuments.doc");
    dstDoc->Save(dataDir);

    std::cout << "\nDocument appended successfully.\nFile saved at" << dataDir.ToUtf8String() << '\n';
}
