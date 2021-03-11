#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/ImportFormatOptions.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>

using namespace Aspose::Words;

void AppendWithImportFormatOptions()
{
    System::String inputDataDir = GetInputDataDir_JoiningAndAppending();
    System::String outputDataDir = GetOutputDataDir_JoiningAndAppending();

    //ExStart:AppendWithImportFormatOptions
    auto srcDoc = System::MakeObject<Document>(inputDataDir + u"Document source with list.docx");
    auto dstDoc = System::MakeObject<Document>(inputDataDir + u"Document destination with list.docx");

    // Specify that if numbering clashes in source and destination documents,
    // then numbering from the source document will be used.
    auto options = System::MakeObject<ImportFormatOptions>();
    options->set_KeepSourceNumbering(true);

    dstDoc->AppendDocument(srcDoc, ImportFormatMode::UseDestinationStyles, options);
    //ExEnd:AppendWithImportFormatOptions
}
