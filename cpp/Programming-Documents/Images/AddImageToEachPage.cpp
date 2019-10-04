#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Layout/Public/LayoutCollector.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeList.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Layout;

namespace
{
    typedef System::Collections::Generic::IEnumerable<System::SharedPtr<Node>> TNodeEnumerable;

    void AddImageToPage(System::SharedPtr<Paragraph> para, int32_t page, System::String const &dataDir)
    {
        System::SharedPtr<Document> doc = System::DynamicCast<Aspose::Words::Document>(para->get_Document());

        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        builder->MoveTo(para);

        // Add a logo to the top left of the page. The image is placed in front of all other text.
        System::SharedPtr<Shape> shape = builder->InsertImage(dataDir + u"Aspose Logo.png", RelativeHorizontalPosition::Page, 60.0, RelativeVerticalPosition::Page, 60.0, -1.0, -1.0, WrapType::None);

        // Add a textbox next to the image which contains some text consisting of the page number. 
        System::SharedPtr<Shape> textBox = System::MakeObject<Shape>(doc, ShapeType::TextBox);

        // We want a floating shape relative to the page.
        textBox->set_WrapType(WrapType::None);
        textBox->set_RelativeHorizontalPosition(RelativeHorizontalPosition::Page);
        textBox->set_RelativeVerticalPosition(RelativeVerticalPosition::Page);

        // Set the textbox position.
        textBox->set_Height(30);
        textBox->set_Width(200);
        textBox->set_Left(150);
        textBox->set_Top(80);

        // Add the textbox and set text.
        textBox->AppendChild(System::MakeObject<Paragraph>(doc));
        builder->InsertNode(textBox);
        builder->MoveTo(textBox->get_FirstChild());
        builder->Writeln(System::String(u"This is a custom note for page ") + page);
    }
}

void AddImageToEachPage()
{
    std::cout << "AddImageToEachPage example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetInputDataDir_WorkingWithImages();
    System::String filename = u"TestFile.doc";
    // This a document that we want to add an image and custom text for each page without using the header or footer.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + filename);

    // Create and attach collector before the document before page layout is built.
    System::SharedPtr<LayoutCollector> layoutCollector = System::MakeObject<LayoutCollector>(doc);

    // Images in a document are added to paragraphs, so to add an image to every page we need to find at any paragraph 
    // Belonging to each page.
    auto enumerator = doc->SelectNodes(u"//Paragraph")->GetEnumerator();

    // Loop through each document page.
    for (int32_t page = 1; page <= doc->get_PageCount(); page++)
    {
        while (enumerator->MoveNext())
        {
            // Check if the current paragraph belongs to the target page.
            System::SharedPtr<Paragraph> paragraph = System::DynamicCast<Paragraph>(enumerator->get_Current());
            if (layoutCollector->GetStartPageIndex(paragraph) == page)
            {
                AddImageToPage(paragraph, page, dataDir);
                break;
            }
        }
    }

    // Call UpdatePageLayout() method if file is to be saved as PDF or image format
    doc->UpdatePageLayout();

    System::String outputPath = GetOutputDataDir_WorkingWithFields() + filename;
    doc->Save(outputPath);

    std::cout << "Inserted images on each page of the document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "AddImageToEachPage example finished." << std::endl << std::endl;
}