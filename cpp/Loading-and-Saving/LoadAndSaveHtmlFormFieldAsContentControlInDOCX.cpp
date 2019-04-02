#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/HtmlLoadOptions.h>

using namespace Aspose::Words;

void LoadAndSaveHtmlFormFieldAsContentControlInDOCX()
{
    std::cout << "LoadAndSaveHtmlFormFieldAsContentControlInDOCX example started." << std::endl;
    // ExStart:LoadAndSaveHtmlFormFieldasContentControlinDOCX
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();

    System::SharedPtr<HtmlLoadOptions> lo = System::MakeObject<HtmlLoadOptions>();
    lo->set_PreferredControlType(HtmlControlType::StructuredDocumentTag);

    //Load the HTML document
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"input.html", lo);

    //Save the HTML document into DOCX
    doc->Save(dataDir + GetOutputFilePath(u"LoadAndSaveHtmlFormFieldAsContentControlInDOCX.docx"), SaveFormat::Docx);
    // ExEnd:LoadAndSaveHtmlFormFieldasContentControlinDOCX
    std::cout << "Html form fields are exported as content control successfully." << std::endl;
    std::cout << "LoadAndSaveHtmlFormFieldAsContentControlInDOCX example finished." << std::endl << std::endl;
}