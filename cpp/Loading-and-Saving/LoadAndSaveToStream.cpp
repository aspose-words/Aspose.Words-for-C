#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/io/stream.h>
#include <system/io/memory_stream.h>
#include <system/io/file.h>
#include <system/console.h>
#include <Model/Document/SaveFormat.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void LoadAndSaveToStream()
{
    // ExStart:LoadAndSaveToStream 
    // ExStart:OpeningFromStream 
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();
    
    // Open the stream. Read only access is enough for Aspose.Words to load a document.
    System::SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(dataDir + u"Document.doc");
    
    // Load the entire document into memory.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(stream);
    
    // You can close the stream now, it is no longer needed because the document is in memory.
    stream->Close();
    // ExEnd:OpeningFromStream 
    
    // ... do something with the document
    
    // Convert the document to a different format and save to stream.
    System::SharedPtr<System::IO::MemoryStream> dstStream = System::MakeObject<System::IO::MemoryStream>();
    doc->Save(dstStream, Aspose::Words::SaveFormat::Doc);
    
    // Rewind the stream position back to zero so it is ready for the next reader.
    dstStream->set_Position(0);
    // ExEnd:LoadAndSaveToStream 
    // Save the document from stream, to disk. Normally you would do something with the stream directly,
    // For example writing the data to a database.
    dataDir = dataDir + GetOutputFilePath(u"LoadAndSaveToStream.doc");
    System::IO::File::WriteAllBytes(dataDir, dstStream->ToArray());
    
    std::cout << "\nStream of document saved successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}