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

void InsertTextInputFormField(System::String dataDir)
{
    // ExStart:DocumentBuilderInsertTextInputFormField
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    builder->InsertTextInput(u"TextInput", Aspose::Words::Fields::TextFormFieldType::Regular, u"", u"Hello", 0);
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderInsertElements.InsertTextInputFormField.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderInsertTextInputFormField
    std::cout << "\nText input form field using DocumentBuilder inserted successfully into a document.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void InsertCheckBoxFormField(System::String dataDir)
{
    // ExStart:DocumentBuilderInsertCheckBoxFormField
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    builder->InsertCheckBox(u"CheckBox", true, true, 0);
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderInsertElements.InsertCheckBoxFormField.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderInsertCheckBoxFormField
    std::cout << "\nCheckbox form field using DocumentBuilder inserted successfully into a document.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void InsertComboBoxFormField(System::String dataDir)
{
    // ExStart:DocumentBuilderInsertComboBoxFormField
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    System::ArrayPtr<System::String> items = System::MakeArray<System::String>({u"One", u"Two", u"Three"});
    builder->InsertComboBox(u"DropDown", items, 0);
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderInsertElements.InsertComboBoxFormField.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderInsertComboBoxFormField
    std::cout << "\nCombobox form field using DocumentBuilder inserted successfully into a document.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void InsertHyperlink(System::String dataDir)
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
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderInsertElements.InsertHyperlink.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderInsertHyperlink
    std::cout << "\nHyperlink using DocumentBuilder inserted successfully into a document.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void InsertOleObject(System::String dataDir)
{
    // ExStart:DocumentBuilderInsertOleObject
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->InsertOleObject(u"http://www.aspose.com", u"htmlfile", true, true, nullptr);
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderInsertElements.InsertOleObject.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderInsertOleObject
    std::cout << "\nOleObject using DocumentBuilder inserted successfully into a document.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void InsertOleObjectwithOlePackage(System::String dataDir)
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
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderInsertElements.InsertOleObjectwithOlePackage.doc");
    doc->Save(dataDir);

    // ExEnd:InsertOleObjectwithOlePackage
    std::cout << "\nOleObject using DocumentBuilder inserted successfully into a document.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

}

void DocumentBuilderInsertElements()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    InsertTextInputFormField(dataDir);
    InsertCheckBoxFormField(dataDir);
    InsertComboBoxFormField(dataDir);
    InsertHyperlink(dataDir);
    InsertOleObject(dataDir);
    InsertOleObjectwithOlePackage(dataDir);
}
