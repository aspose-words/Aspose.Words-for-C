#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/HtmlLoadOptions.h>

using namespace Aspose::Words;

void LoadAndSaveHtmlFormFieldAsContentControlInDOCX()
{
    std::cout << "LoadAndSaveHtmlFormFieldAsContentControlInDOCX example started." << std::endl;
    // ExStart:LoadAndSaveHtmlFormFieldasContentControlinDOCX
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

    System::SharedPtr<HtmlLoadOptions> lo = System::MakeObject<HtmlLoadOptions>();
    lo->set_PreferredControlType(HtmlControlType::StructuredDocumentTag);

    //Load the HTML document
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"input.html", lo);

    //Save the HTML document into DOCX
    doc->Save(outputDataDir + u"LoadAndSaveHtmlFormFieldAsContentControlInDOCX.docx", SaveFormat::Docx);
    // ExEnd:LoadAndSaveHtmlFormFieldasContentControlinDOCX
    std::cout << "Html form fields are exported as content control successfully." << std::endl;
    std::cout << "LoadAndSaveHtmlFormFieldAsContentControlInDOCX example finished." << std::endl << std::endl;
}