#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/console.h>
#include <system/collections/list.h>
#include <system/collections/ienumerator.h>
#include <system/array.h>
#include <Model/Sections/SectionCollection.h>
#include <Model/Sections/Section.h>
#include <Model/Nodes/Node.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

namespace
{
    void DoPrepend(const System::SharedPtr<Document>& dstDoc, const System::SharedPtr<Document>& srcDoc, ImportFormatMode mode)
    {
        // Loop through all sections in the source document. 
        // Section nodes are immediate children of the Document node so we can just enumerate the Document.
        auto sections = srcDoc->get_Sections()->ToArray();

        // Reverse the order of the sections so they are prepended to start of the destination document in the correct order.
        for (auto it = sections->data().rbegin(); it != sections->data().rend(); ++it)
        {
            // Import the nodes from the source document.
            System::SharedPtr<Node> dstSection = dstDoc->ImportNode(*it, true, mode);

            // Now the new section node can be prepended to the destination document.
            // Note how PrependChild is used instead of AppendChild. This is the only line changed compared 
            // To the original method.
            dstDoc->PrependChild(dstSection);
        }
    }
}

void PrependDocument()
{
    std::cout << "PrependDocument example started." << std::endl;
    // ExStart:PrependDocument
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");

    // Append the source document to the destination document. This causes the result to have line spacing problems.
    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

    // Instead prepend the content of the destination document to the start of the source document.
    // This results in the same joined document but with no line spacing issues.
    DoPrepend(srcDoc, dstDoc, ImportFormatMode::KeepSourceFormatting);

    System::String outputPath = dataDir + GetOutputFilePath(u"PrependDocument.doc");
    // Save the document
    dstDoc->Save(outputPath);
    // ExEnd:PrependDocument
    std::cout << "Document prepended successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "PrependDocument example finished." << std::endl << std::endl;
}
