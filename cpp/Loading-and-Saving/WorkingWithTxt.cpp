#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <Model/Document/Document.h>
#include <Model/Document/TxtLeadingSpacesOptions.h>
#include <Model/Document/TxtLoadOptions.h>
#include <Model/Document/TxtTrailingSpacesOptions.h>
#include <Model/Saving/TxtSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void SaveAsTxt(System::String const &dataDir)
    {
        //ExStart:SaveAsTxt
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Document.doc");
        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingwithTxt.SaveAsTxt.txt");
        doc->Save(outputPath);
        //ExEnd:SaveAsTxt
        std::cout << "Document saved as TXT." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void AddBidiMarks(System::String const &dataDir)
    {
        //ExStart:AddBidiMarks
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Input.docx");
        System::SharedPtr<TxtSaveOptions> saveOptions = System::MakeObject<TxtSaveOptions>();
        saveOptions->set_AddBidiMarks(false);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingwithTxt.AddBidiMarks.txt");
        doc->Save(outputPath, saveOptions);
        //ExEnd:AddBidiMarks
        std::cout << "Add bi-directional marks set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void DetectNumberingWithWhitespaces(System::String const &dataDir)
    {
        //ExStart:DetectNumberingWithWhitespaces
        System::SharedPtr<TxtLoadOptions> loadOptions = System::MakeObject<TxtLoadOptions>();
        loadOptions->set_DetectNumberingWithWhitespaces(false);

        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"LoadTxt.txt", loadOptions);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingwithTxt.DetectNumberingWithWhitespaces.docx");
        doc->Save(outputPath);
        //ExEnd:DetectNumberingWithWhitespaces
        std::cout << "Detect number with whitespaces successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void HandleSpacesOptions(System::String const &dataDir)
    {
        //ExStart:HandleSpacesOptions
        System::SharedPtr<TxtLoadOptions> loadOptions = System::MakeObject<TxtLoadOptions>();

        loadOptions->set_LeadingSpacesOptions(TxtLeadingSpacesOptions::Trim);
        loadOptions->set_TrailingSpacesOptions(TxtTrailingSpacesOptions::Trim);
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"LoadTxt.txt", loadOptions);

        System::String outputPath = dataDir + GetOutputFilePath(u"WorkingwithTxt.HandleSpacesOptions.docx");
        doc->Save(outputPath);
        //ExEnd:HandleSpacesOptions
        std::cout << "Trim leading and trailing spaces while importing text document." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkingWithTxt()
{
    std::cout << "WorkingWithTxt example started." << std::endl;
    //ExStart:WorkingWithTxt
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();

    SaveAsTxt(dataDir);
    AddBidiMarks(dataDir);
    DetectNumberingWithWhitespaces(dataDir);
    HandleSpacesOptions(dataDir);
    //ExEnd:WorkingWithTxt
    std::cout << "WorkingWithTxt example finished." << std::endl << std::endl;
}