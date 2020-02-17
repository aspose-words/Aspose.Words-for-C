#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OleFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Ole/OlePackage.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Underline.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;

namespace
{
    void InsertTextInputFormField(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderInsertTextInputFormField
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u"Hello", 0);
        System::String outputPath = outputDataDir + u"DocumentBuilderInsertElements.InsertTextInputFormField.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertTextInputFormField
        std::cout << "Text input form field using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertCheckBoxFormField(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderInsertCheckBoxFormField
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->InsertCheckBox(u"CheckBox", true, true, 0);
        System::String outputPath = outputDataDir + u"DocumentBuilderInsertElements.InsertCheckBoxFormField.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertCheckBoxFormField
        std::cout << "Checkbox form field using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertComboBoxFormField(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderInsertComboBoxFormField
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::ArrayPtr<System::String> items = System::MakeArray<System::String>({u"One", u"Two", u"Three"});
        builder->InsertComboBox(u"DropDown", items, 0);
        System::String outputPath = outputDataDir + u"DocumentBuilderInsertElements.InsertComboBoxFormField.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertComboBoxFormField
        std::cout << "Combobox form field using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertHtml(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderInsertHtml
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->InsertHtml(u"<P align='right'>Paragraph right</P><b>Implicit paragraph left</b><div align='center'>Div center</div><h1 align='left'>Heading 1 left.</h1>");
        System::String outputPath = outputDataDir + u"DocumentBuilderInsertElements.InsertHtml.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertHtml
        std::cout << "HTML using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;

    }

    void InsertHyperlink(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderInsertHyperlink
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Please make sure to visit ");

        // Specify font formatting for the hyperlink.
        builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
        builder->get_Font()->set_Underline(Underline::Single);
        // Insert the link.
        builder->InsertHyperlink(u"Aspose Website", u"http://www.aspose.com", false);

        // Revert to default formatting.
        builder->get_Font()->ClearFormatting();

        builder->Write(u" for more information.");
        System::String outputPath = outputDataDir + u"DocumentBuilderInsertElements.InsertHyperlink.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertHyperlink
        std::cout << "Hyperlink using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertTableOfContents(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderInsertTableOfContents
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Insert a table of contents at the beginning of the document.
        builder->InsertTableOfContents(u"\\o \"1-3\" \\h \\z \\u");

        // Start the actual document content on the second page.
        builder->InsertBreak(BreakType::PageBreak);

        // Build a document with complex structure by applying different heading styles thus creating TOC entries.
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);

        builder->Writeln(u"Heading 1");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

        builder->Writeln(u"Heading 1.1");
        builder->Writeln(u"Heading 1.2");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);

        builder->Writeln(u"Heading 2");
        builder->Writeln(u"Heading 3");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

        builder->Writeln(u"Heading 3.1");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading3);

        builder->Writeln(u"Heading 3.1.1");
        builder->Writeln(u"Heading 3.1.2");
        builder->Writeln(u"Heading 3.1.3");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

        builder->Writeln(u"Heading 3.2");
        builder->Writeln(u"Heading 3.3");

        doc->UpdateFields();
        System::String outputPath = outputDataDir + u"DocumentBuilderInsertElements.InsertTableOfContents.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertTableOfContents
        std::cout << "Table of contents using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertOleObject(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderInsertOleObject
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        builder->InsertOleObject(u"http://www.aspose.com", u"htmlfile", true, true, nullptr);
        System::String outputPath = outputDataDir + u"DocumentBuilderInsertElements.InsertOleObject.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertOleObject
        std::cout << "OleObject using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertOleObjectwithOlePackage(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:InsertOleObjectwithOlePackage
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        System::ArrayPtr<uint8_t> bs = System::IO::File::ReadAllBytes(inputDataDir + u"input.zip");

        System::SharedPtr<System::IO::Stream> stream = System::MakeObject<System::IO::MemoryStream>(bs);
        System::SharedPtr<Shape> shape = builder->InsertOleObject(stream, u"Package", true, nullptr);
        System::SharedPtr<OlePackage> olePackage = shape->get_OleFormat()->get_OlePackage();
        olePackage->set_FileName(u"filename.zip");
        olePackage->set_DisplayName(u"displayname.zip");
        System::String outputPath = outputDataDir + u"DocumentBuilderInsertElements.InsertOleObjectwithOlePackage.doc";
        doc->Save(outputPath);
        // ExEnd:InsertOleObjectwithOlePackage
        std::cout << "OleObject using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void GetAccessToOLEObjectRawData(System::String const &inputDataDir)
    {
        // ExStart:GetAccessToOLEObjectRawData
        // Load document with OLE object.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"DocumentBuilderInsertOleObject.doc");
        System::SharedPtr<Shape> oleShape = System::StaticCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        System::ByteArrayPtr oleRawData = oleShape->get_OleFormat()->GetRawData();
        // ExEnd:GetAccessToOLEObjectRawData
    }
}

void DocumentBuilderInsertElements()
{
    std::cout << "DocumentBuilderInsertElements example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    InsertTextInputFormField(outputDataDir);
    InsertCheckBoxFormField(outputDataDir);
    InsertComboBoxFormField(outputDataDir);
    InsertHtml(outputDataDir);
    InsertHyperlink(outputDataDir);
    InsertTableOfContents(outputDataDir);
    InsertOleObject(outputDataDir);
    InsertOleObjectwithOlePackage(inputDataDir, outputDataDir);
    GetAccessToOLEObjectRawData(inputDataDir);
    std::cout << "DocumentBuilderInsertElements example finished." << std::endl << std::endl;
}
