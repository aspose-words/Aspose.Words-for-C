#include "stdafx.h"
#include "examples.h"

#include <system/io/file.h>
#include <system/io/memory_stream.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void LoadAndSaveToStream()
{
    std::cout << "LoadAndSaveToStream example started." << std::endl;
    // ExStart:LoadAndSaveToStream
    // ExStart:OpeningFromStream
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

    // Open the stream. Read only access is enough for Aspose.Words to load a document.
    System::SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(inputDataDir + u"Document.docx");

    // Load the entire document into memory.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(stream);

    // You can close the stream now, it is no longer needed because the document is in memory.
    stream->Close();
    // ExEnd:OpeningFromStream

    // ... do something with the document

    // Convert the document to a different format and save to stream.
    System::SharedPtr<System::IO::MemoryStream> dstStream = System::MakeObject<System::IO::MemoryStream>();
    doc->Save(dstStream, SaveFormat::Doc);

    // Rewind the stream position back to zero so it is ready for the next reader.
    dstStream->set_Position(0);
    // Save the document from stream, to disk. Normally you would do something with the stream directly,
    // For example writing the data to a database.
    System::String outputPath = outputDataDir + u"LoadAndSaveToStream.docx";
    System::IO::File::WriteAllBytes(outputPath, dstStream->ToArray());
    // ExEnd:LoadAndSaveToStream
    std::cout << "Stream of document saved successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "LoadAndSaveToStream example finished." << std::endl << std::endl;
}