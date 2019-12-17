#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

namespace
{
    void DoPrepend(const System::SharedPtr<Document>& dstDoc, const System::SharedPtr<Document>& srcDoc, ImportFormatMode mode)
    {
        // Loop through all sections in the source document. 
        // Section nodes are immediate children of the Document node so we can just enumerate the Document.
        System::ArrayPtr<System::SharedPtr<Section>> sections = srcDoc->get_Sections()->ToArray();

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
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_JoiningAndAppending();
    System::String outputDataDir = GetOutputDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(inputDataDir + u"TestFile.Source.doc");

    // Append the source document to the destination document. This causes the result to have line spacing problems.
    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

    // Instead prepend the content of the destination document to the start of the source document.
    // This results in the same joined document but with no line spacing issues.
    DoPrepend(srcDoc, dstDoc, ImportFormatMode::KeepSourceFormatting);

    System::String outputPath = outputDataDir + u"PrependDocument.doc";
    // Save the document
    dstDoc->Save(outputPath);
    std::cout << "Document prepended successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "PrependDocument example finished." << std::endl << std::endl;
}
