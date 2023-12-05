#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples {
    namespace Getting_started {

        class HelloWorld : public DocsExamplesBase
        {
        public:
            void SimpleHelloWorld()
            {
                //ExStart:HelloWorld
                //GistId:ab3973c31f74954b9fff03abc4ca6d5b
                auto docA = MakeObject<Document>();
                auto builder = MakeObject<DocumentBuilder>(docA);

                // Insert text to the document start.
                builder->MoveToDocumentStart();
                builder->Write(u"First Hello World paragraph");

                auto docB = MakeObject<Document>(MyDir + u"Document.docx");
                // Add document B to the and of document A, preserving document B formatting.
                docA->AppendDocument(docB, ImportFormatMode::KeepSourceFormatting);

                docA->Save(ArtifactsDir + u"HelloWorld.SimpleHelloWorld.pdf");
                //ExEnd:HelloWorld
            }
        };
    }
} // namespace DocsExamples::Programming_with_Documents
