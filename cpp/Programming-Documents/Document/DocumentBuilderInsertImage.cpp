#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Drawing/WrapType.h>
#include <Model/Drawing/Shape.h>
#include <Model/Drawing/RelativeVerticalPosition.h>
#include <Model/Drawing/RelativeHorizontalPosition.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace
{
    void InsertInlineImage(System::String const &dataDir)
    {
        // ExStart:DocumentBuilderInsertInlineImage
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->InsertImage(dataDir + u"Watermark.png");
        System::String outputPath = dataDir + GetOutputFilePath(u"DocumentBuilderInsertImage.InsertInlineImage.doc");
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertInlineImage
        std::cout << "Inline image using DocumentBuilder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertFloatingImage(System::String const &dataDir)
    {
        // ExStart:DocumentBuilderInsertFloatingImage
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->InsertImage(dataDir + u"Watermark.png", RelativeHorizontalPosition::Margin, 100, RelativeVerticalPosition::Margin, 100, 200, 100, WrapType::Square);
        System::String outputPath = dataDir + GetOutputFilePath(u"DocumentBuilderInsertImage.InsertFloatingImage.doc");
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertFloatingImage
        std::cout << "Inline image using DocumentBuilder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void DocumentBuilderInsertImage()
{
    std::cout << "DocumentBuilderInsertImage example started." << std::endl;
     // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    InsertInlineImage(dataDir);
    InsertFloatingImage(dataDir);
    std::cout << "DocumentBuilderInsertImage example finished." << std::endl << std::endl;
}
