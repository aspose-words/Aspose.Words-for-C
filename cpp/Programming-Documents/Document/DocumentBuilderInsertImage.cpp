#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Saving/SaveOutputParameters.h>
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

void InsertInlineImage(System::String dataDir)
{
    // ExStart:DocumentBuilderInsertInlineImage
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    builder->InsertImage(dataDir + u"Watermark.png");
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderInsertImage.InsertInlineImage.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderInsertInlineImage
    std::cout << "\nInline image using DocumentBuilder inserted successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void InsertFloatingImage(System::String dataDir)
{
    // ExStart:DocumentBuilderInsertFloatingImage
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    builder->InsertImage(dataDir + u"Watermark.png", Aspose::Words::Drawing::RelativeHorizontalPosition::Margin, 100, Aspose::Words::Drawing::RelativeVerticalPosition::Margin, 100, 200, 100, Aspose::Words::Drawing::WrapType::Square);
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderInsertImage.InsertFloatingImage.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderInsertFloatingImage
    std::cout << "\nInline image using DocumentBuilder inserted successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
}

void DocumentBuilderInsertImage()
{
    
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    InsertInlineImage(dataDir);
    InsertFloatingImage(dataDir);
}
