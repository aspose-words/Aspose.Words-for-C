#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Fields/FormField.h>
#include <Aspose.Words.Cpp/Fields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Range.h>
#include <drawing/color.h>
#include <system/array.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithFormFields : public DocsExamplesBase
{
public:
    void InsertFormFields()
    {
        //ExStart:InsertFormFields
        //GistId:b0f2e2ef2b8f321b68cd22b92f8f549b
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        ArrayPtr<String> items = MakeArray<String>({u"One", u"Two", u"Three"});
        builder->InsertComboBox(u"DropDown", items, 0);
        //ExEnd:InsertFormFields
    }

    void FormFieldsWorkWithProperties()
    {
        //ExStart:FormFieldsWorkWithProperties
        //GistId:b0f2e2ef2b8f321b68cd22b92f8f549b
        auto doc = MakeObject<Document>(MyDir + u"Form fields.docx");
        SharedPtr<FormField> formField = doc->get_Range()->get_FormFields()->idx_get(3);

        if (formField->get_Type() == FieldType::FieldFormTextInput)
        {
            formField->set_Result(String(u"My name is ") + formField->get_Name());
        }
        //ExEnd:FormFieldsWorkWithProperties
    }

    void FormFieldsGetFormFieldsCollection()
    {
        //ExStart:FormFieldsGetFormFieldsCollection
        //GistId:b0f2e2ef2b8f321b68cd22b92f8f549b
        auto doc = MakeObject<Document>(MyDir + u"Form fields.docx");

        SharedPtr<FormFieldCollection> formFields = doc->get_Range()->get_FormFields();
        //ExEnd:FormFieldsGetFormFieldsCollection
    }

    void FormFieldsGetByName()
    {
        //ExStart:FormFieldsFontFormatting
        //GistId:b0f2e2ef2b8f321b68cd22b92f8f549b
        //ExStart:FormFieldsGetByName
        //GistId:b0f2e2ef2b8f321b68cd22b92f8f549b
        auto doc = MakeObject<Document>(MyDir + u"Form fields.docx");

        SharedPtr<FormFieldCollection> documentFormFields = doc->get_Range()->get_FormFields();

        SharedPtr<FormField> formField1 = documentFormFields->idx_get(3);
        SharedPtr<FormField> formField2 = documentFormFields->idx_get(u"Text2");
        //ExEnd:FormFieldsGetByName

        formField1->get_Font()->set_Size(20);
        formField2->get_Font()->set_Color(System::Drawing::Color::get_Red());
        //ExEnd:FormFieldsFontFormatting
    }
};

}} // namespace DocsExamples::Programming_with_Documents
