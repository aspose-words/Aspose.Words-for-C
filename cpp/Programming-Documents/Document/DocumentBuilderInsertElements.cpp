#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/io/stream.h>
#include <system/io/memory_stream.h>
#include <system/io/file.h>
#include <system/details/dispose_guard.h>
#include <system/array.h>
#include <Model/Text/Underline.h>
#include <Model/Text/Font.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Fields/FormFields/TextFormFieldType.h>
#include <Model/Fields/FormFields/FormField.h>
#include <Model/Fields/Field.h>
#include <Model/Drawing/Shape.h>
#include <Model/Drawing/Ole/OlePackage.h>
#include <Model/Drawing/Ole/OleFormat.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <drawing/color.h>
#include <cstdint>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;

namespace
{
    void InsertTextInputFormField(System::String const &dataDir)
    {
        // ExStart:DocumentBuilderInsertTextInputFormField
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->InsertTextInput(u"TextInput", Aspose::Words::Fields::TextFormFieldType::Regular, u"", u"Hello", 0);
        System::String outputPath = dataDir + GetOutputFilePath(u"DocumentBuilderInsertElements.InsertTextInputFormField.doc");
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertTextInputFormField
        std::cout << "Text input form field using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertCheckBoxFormField(System::String const &dataDir)
    {
        // ExStart:DocumentBuilderInsertCheckBoxFormField
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->InsertCheckBox(u"CheckBox", true, true, 0);
        System::String outputPath = dataDir + GetOutputFilePath(u"DocumentBuilderInsertElements.InsertCheckBoxFormField.doc");
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertCheckBoxFormField
        std::cout << "Checkbox form field using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertComboBoxFormField(System::String const &dataDir)
    {
        // ExStart:DocumentBuilderInsertComboBoxFormField
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::ArrayPtr<System::String> items = System::MakeArray<System::String>({u"One", u"Two", u"Three"});
        builder->InsertComboBox(u"DropDown", items, 0);
        System::String outputPath = dataDir + GetOutputFilePath(u"DocumentBuilderInsertElements.InsertComboBoxFormField.doc");
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertComboBoxFormField
        std::cout << "Combobox form field using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertHyperlink(System::String const &dataDir)
    {
        // ExStart:DocumentBuilderInsertHyperlink
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Please make sure to visit ");

        // Specify font formatting for the hyperlink.
        builder->get_Font()->set_Color(System::Drawing::Color::get_Blue());
        builder->get_Font()->set_Underline(Aspose::Words::Underline::Single);
        // Insert the link.
        builder->InsertHyperlink(u"Aspose Website", u"http://www.aspose.com", false);

        // Revert to default formatting.
        builder->get_Font()->ClearFormatting();

        builder->Write(u" for more information.");
        System::String outputPath = dataDir + GetOutputFilePath(u"DocumentBuilderInsertElements.InsertHyperlink.doc");
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertHyperlink
        std::cout << "Hyperlink using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertOleObject(System::String const &dataDir)
    {
        // ExStart:DocumentBuilderInsertOleObject
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        builder->InsertOleObject(u"http://www.aspose.com", u"htmlfile", true, true, nullptr);
        System::String outputPath = dataDir + GetOutputFilePath(u"DocumentBuilderInsertElements.InsertOleObject.doc");
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertOleObject
        std::cout << "OleObject using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertOleObjectwithOlePackage(System::String const &dataDir)
    {
        // ExStart:InsertOleObjectwithOlePackage
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        System::ArrayPtr<uint8_t> bs = System::IO::File::ReadAllBytes(dataDir + u"input.zip");

        System::SharedPtr<System::IO::Stream> stream = System::MakeObject<System::IO::MemoryStream>(bs);
        System::SharedPtr<Shape> shape = builder->InsertOleObject(stream, u"Package", true, nullptr);
        System::SharedPtr<OlePackage> olePackage = shape->get_OleFormat()->get_OlePackage();
        olePackage->set_FileName(u"filename.zip");
        olePackage->set_DisplayName(u"displayname.zip");
        System::String outputPath = dataDir + GetOutputFilePath(u"DocumentBuilderInsertElements.InsertOleObjectwithOlePackage.doc");
        doc->Save(outputPath);
        // ExEnd:InsertOleObjectwithOlePackage
        std::cout << "OleObject using DocumentBuilder inserted successfully into a document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void DocumentBuilderInsertElements()
{
    std::cout << "DocumentBuilderInsertElements example started." << std::endl;
    // ExStart:DocumentBuilderInsertElements
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    InsertTextInputFormField(dataDir);
    InsertCheckBoxFormField(dataDir);
    InsertComboBoxFormField(dataDir);
    InsertHyperlink(dataDir);
    InsertOleObject(dataDir);
    InsertOleObjectwithOlePackage(dataDir);
    // ExEnd:DocumentBuilderInsertElements
    std::cout << "DocumentBuilderInsertElements example finished." << std::endl << std::endl;
}
