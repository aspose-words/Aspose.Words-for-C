#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>

#include "MailMergeCommon.h"

using namespace Aspose::Words;

void SimpleMailMerge()
{
    std::cout << "SimpleMailMerge example started." << std::endl;
    // ExStart:SimpleMailMerge
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_MailMergeAndReporting();
    System::String outputDataDir = GetOutputDataDir_MailMergeAndReporting();

	// Include the code for our template.
    auto doc = System::MakeObject<Document>();
	auto builder = System::MakeObject<DocumentBuilder>(doc);


    // Create Merge Fields.
    builder->InsertField(u" MERGEFIELD CustomerName ");
    builder->InsertParagraph();
    builder->InsertField(u" MERGEFIELD Item ");
    builder->InsertParagraph();
    builder->InsertField(u" MERGEFIELD Quantity ");

	doc->Save(outputDataDir + u"MailMerge.TestTemplate.docx");
	
    // Fill the fields in the document with user data.
    doc->get_MailMerge()->Execute(
		System::MakeArray<System::String>({ u"CustomerName", u"Item", u"Quantity" }),
        Box<System::String, System::String, System::String>( u"John Doe", u"Hawaiian", u"2" )
    );

    doc->Save(outputDataDir + u"MailMerge.SimpleMailMerge.docx");
    // ExEnd:SimpleMailMerge
    std::cout << "Simple Mail merge performed with array data successfully.\n";
    std::cout << "SimpleMailMerge example finished.\n\n";
}