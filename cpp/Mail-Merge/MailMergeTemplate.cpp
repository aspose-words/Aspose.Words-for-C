#include "stdafx.h"
#include "examples.h"
#include "MailMergeCommon.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;

//ExStart:CreateMailMergeTemplate
System::SharedPtr<Aspose::Words::Document> CreateMailMergeTemplate()
{
	auto doc = System::MakeObject<Document>();
	auto builder = System::MakeObject<DocumentBuilder>(doc);
	
	// Insert a text input field the unique name of this field is "Hello", the other parameters define
	// what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
	builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u"Hello", 0);
	builder->InsertField(uR"(MERGEFIELD CustomerFirstName \* MERGEFORMAT)");

	builder->InsertTextInput(u"TextInput1", TextFormFieldType::Regular, u"", u" ", 0);
	builder->InsertField(uR"(MERGEFIELD CustomerLastName \* MERGEFORMAT)");

	builder->InsertTextInput(u"TextInput1", TextFormFieldType::Regular, u"", u" , ", 0);

	// Inserts a paragraph break into the document
	builder->InsertParagraph();

	// Insert mail body
	builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u"Thanks for purchasing our ", 0);
	builder->InsertField(uR"(MERGEFIELD ProductName \* MERGEFORMAT)");

	builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u", please download your Invoice at ",
		0);
	builder->InsertField(uR"(MERGEFIELD InvoiceURL \* MERGEFORMAT)");

	builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"",
		u". If you have any questions please call ", 0);
	builder->InsertField(uR"(MERGEFIELD Supportphone \* MERGEFORMAT)");

	builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u", or email us at ", 0);
	builder->InsertField(uR"(MERGEFIELD SupportEmail \* MERGEFORMAT)");

	builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u".", 0);

	builder->InsertParagraph();

	// Insert mail ending
	builder->InsertTextInput(u"TextInput", TextFormFieldType::Regular, u"", u"Best regards,", 0);
	builder->InsertBreak(BreakType::LineBreak);
	builder->InsertField(uR"(MERGEFIELD EmployeeFullname \* MERGEFORMAT)");

	builder->InsertTextInput(u"TextInput1", TextFormFieldType::Regular, u"", u" ", 0);
	builder->InsertField(uR"(MERGEFIELD EmployeeDepartment \* MERGEFORMAT)");

	return doc;
}
//ExEnd:CreateMailMergeTemplate
