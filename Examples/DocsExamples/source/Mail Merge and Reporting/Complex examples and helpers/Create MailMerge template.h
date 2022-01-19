#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FormField.h>
#include <Aspose.Words.Cpp/Fields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <system/string.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace DocsExamples { namespace Mail_Merge_and_Reporting { namespace Custom_examples {

class CreateMailMergeTemplate : public System::Object
{
public:
    //ExStart:CreateMailMergeTemplate
    SharedPtr<Document> Template()
    {
        auto builder = MakeObject<DocumentBuilder>();

        // Insert a text input field the unique name of this field is "Hello", the other parameters define
        // what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
        builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u"Hello", 0);
        builder->InsertField(u"MERGEFIELD CustomerFirstName \\* MERGEFORMAT");

        builder->InsertTextInput(u"TextInput1", TextFormFieldType::Regular, u"", u" ", 0);
        builder->InsertField(u"MERGEFIELD CustomerLastName \\* MERGEFORMAT");

        builder->InsertTextInput(u"TextInput1", TextFormFieldType::Regular, u"", u" , ", 0);

        // Inserts a paragraph break into the document
        builder->InsertParagraph();

        // Insert mail body
        builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u"Thanks for purchasing our ", 0);
        builder->InsertField(u"MERGEFIELD ProductName \\* MERGEFORMAT");

        builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u", please download your Invoice at ", 0);
        builder->InsertField(u"MERGEFIELD InvoiceURL \\* MERGEFORMAT");

        builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u". If you have any questions please call ", 0);
        builder->InsertField(u"MERGEFIELD Supportphone \\* MERGEFORMAT");

        builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u", or email us at ", 0);
        builder->InsertField(u"MERGEFIELD SupportEmail \\* MERGEFORMAT");

        builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u".", 0);

        builder->InsertParagraph();

        // Insert mail ending
        builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u"Best regards,", 0);
        builder->InsertBreak(BreakType::LineBreak);
        builder->InsertField(u"MERGEFIELD EmployeeFullname \\* MERGEFORMAT");

        builder->InsertTextInput(u"TextInput1", TextFormFieldType::Regular, u"", u" ", 0);
        builder->InsertField(u"MERGEFIELD EmployeeDepartment \\* MERGEFORMAT");

        return builder->get_Document();
    }
    //ExEnd:CreateMailMergeTemplate
};

}}} // namespace DocsExamples::Mail_Merge_and_Reporting::Custom_examples
