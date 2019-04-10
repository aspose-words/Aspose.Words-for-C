#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Model/Sections/Section.h>
#include <Model/Nodes/Node.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void AppendDocumentManually()
{
    std::cout << "AppendDocumentManually example started." << std::endl;
    // ExStart:AppendDocumentManually
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");
    ImportFormatMode mode = ImportFormatMode::KeepSourceFormatting;

    // Loop through all sections in the source document. 
    // Section nodes are immediate children of the Document node so we can just enumerate the Document.
    for (System::SharedPtr<Section> srcSection : System::IterateOver<System::SharedPtr<Section>>(srcDoc))
    {
        // Because we are copying a section from one document to another, 
        // It is required to import the Section node into the destination document.
        // This adjusts any document-specific references to styles, lists, etc.
        // Importing a node creates a copy of the original node, but the copy
        // Is ready to be inserted into the destination document.
        System::SharedPtr<Node> dstSection = dstDoc->ImportNode(srcSection, true, mode);

        // Now the new section node can be appended to the destination document.
        dstDoc->AppendChild(dstSection);
    }

    System::String outputPath = dataDir + GetOutputFilePath(u"AppendDocumentManually.doc");
    // Save the joined document
    dstDoc->Save(outputPath);
    // ExEnd:AppendDocumentManually
    std::cout << "Document appended successfully with manual append operation." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "AppendDocumentManually example finished." << std::endl << std::endl;
}
