#include "stdafx.h"
#include "examples.h"
#include <system/string.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>


void FormFieldsFontFormatting()
{
	// ExStart:FormFieldsFontFormatting
	const System::String inputDataDir = GetInputDataDir_WorkingWithFields();

	auto doc = System::MakeObject<Aspose::Words::Document>(inputDataDir + u"Document.doc");
	doc->get_Range()->get_FormFields()->idx_get(0)->get_Font()->set_Size(20);
	doc->get_Range()->get_FormFields()->idx_get(0)->get_Font()->set_Color(System::Drawing::Color::get_Red());

	const System::String outDataDir = GetOutputDataDir_WorkingWithFields();
	doc->Save(outDataDir + u"FormFieldsFontFormatting.doc");
	// ExEnd:FormFieldsFontFormatting
}