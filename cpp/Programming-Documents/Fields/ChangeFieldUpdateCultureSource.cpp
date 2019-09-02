#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Fields/FieldOptions.h>
#include <Model/Fields/FieldUpdateCultureSource.h>
#include <Model/MailMerge/MailMerge.h>
#include <Model/Text/Font.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void ChangeFieldUpdateCultureSource()
{
    std::cout << "ChangeFieldUpdateCultureSource example started." << std::endl;
    // ExStart:ChangeFieldUpdateCultureSource
    typedef System::SharedPtr<System::Object> TObjectPtr;
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();
    // We will test this functionality creating a document with two fields with date formatting
    // ExStart:DocumentBuilderInsertField
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Insert content with German locale.
    builder->get_Font()->set_LocaleId(1031);
    builder->InsertField(u"MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
    builder->Write(u" - ");
    builder->InsertField(u"MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");
    // ExEnd:DocumentBuilderInsertField
    // Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from.
    // Set the culture used during field update to the culture used by the field.
    doc->get_FieldOptions()->set_FieldUpdateCultureSource(FieldUpdateCultureSource::FieldCode);
    doc->get_MailMerge()->Execute(System::MakeArray<System::String>({u"Date2"}), System::MakeArray<TObjectPtr>({System::ObjectExt::Box<System::DateTime>(System::DateTime(2011, 1, 1))}));
    System::String outputPath = outputDataDir + u"ChangeFieldUpdateCultureSource.doc";
    doc->Save(outputPath);
    // ExEnd:ChangeFieldUpdateCultureSource

    std::cout << "Culture changed successfully used in formatting fields during update." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ChangeFieldUpdateCultureSource example finished." << std::endl << std::endl;
}