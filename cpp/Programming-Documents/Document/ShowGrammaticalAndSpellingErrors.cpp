#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void ShowGrammaticalAndSpellingErrors()
{
    std::cout << "ShowGrammaticalAndSpellingErrors example started.\n";

	System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
	System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

	// ExStart: ShowGrammaticalAndSpellingErrors
	auto doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
	doc->set_ShowGrammaticalErrors(true);
	doc->set_ShowSpellingErrors(true);

	auto outputPath = outputDataDir + u"ShowGrammaticalAndSpellingErrors.docx";
	doc->Save(outputPath);
	// ExEnd: ShowGrammaticalAndSpellingErrors
	std::cout << "Document saved successfully.\nFile saved at " << outputPath.ToUtf8String() << '\n';

    std::cout << "ShowGrammaticalAndSpellingErrors example finished.\n\n";
}
