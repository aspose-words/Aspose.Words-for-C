#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace
{
    void InsertInlineImage(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderInsertInlineImage
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->InsertImage(inputDataDir + u"Watermark.png");
        System::String outputPath = outputDataDir + u"DocumentBuilderInsertImage.InsertInlineImage.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertInlineImage
        std::cout << "Inline image using DocumentBuilder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertFloatingImage(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderInsertFloatingImage
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->InsertImage(inputDataDir + u"Watermark.png", RelativeHorizontalPosition::Margin, 100, RelativeVerticalPosition::Margin, 100, 200, 100, WrapType::Square);
        System::String outputPath = outputDataDir + u"DocumentBuilderInsertImage.InsertFloatingImage.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertFloatingImage
        std::cout << "Inline image using DocumentBuilder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void DocumentBuilderInsertImage()
{
    std::cout << "DocumentBuilderInsertImage example started." << std::endl;
     // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    InsertInlineImage(inputDataDir, outputDataDir);
    InsertFloatingImage(inputDataDir, outputDataDir);
    std::cout << "DocumentBuilderInsertImage example finished." << std::endl << std::endl;
}
