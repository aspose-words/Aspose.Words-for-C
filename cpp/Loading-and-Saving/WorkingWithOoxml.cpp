#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace
{
    void EncryptDocxWithPassword(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:EncryptDocxWithPassword
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
        System::SharedPtr<OoxmlSaveOptions> ooxmlSaveOptions = System::MakeObject<OoxmlSaveOptions>();
        ooxmlSaveOptions->set_Password(u"password");
        System::String outputPath = outputDataDir + u"WorkingWithOoxml.EncryptDocxWithPassword.docx";
        doc->Save(outputPath, ooxmlSaveOptions);
        //ExEnd:EncryptDocxWithPassword
        std::cout << "The password of document is set using ECMA376 Standard encryption algorithm." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetOOXMLCompliance(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:SetOOXMLCompliance
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

        // Set Word2016 version for document
        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2016);

        //Set the Strict compliance level. 
        System::SharedPtr<OoxmlSaveOptions> ooxmlSaveOptions = System::MakeObject<OoxmlSaveOptions>();
        ooxmlSaveOptions->set_Compliance(OoxmlCompliance::Iso29500_2008_Strict);
        ooxmlSaveOptions->set_SaveFormat(SaveFormat::Docx);
        System::String outputPath = outputDataDir + u"WorkingWithOoxml.SetOOXMLCompliance.docx";
        doc->Save(outputPath, ooxmlSaveOptions);
        //ExEnd:SetOOXMLCompliance
        std::cout << "Document is saved with ISO/IEC 29500:2008 Strict compliance level." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void UpdateLastSavedTimeProperty(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:UpdateLastSavedTimeProperty
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

        System::SharedPtr<OoxmlSaveOptions> ooxmlSaveOptions = System::MakeObject<OoxmlSaveOptions>();
        ooxmlSaveOptions->set_UpdateLastSavedTimeProperty(true);

        System::String outputPath = outputDataDir + u"WorkingWithOoxml.UpdateLastSavedTimeProperty.docx";
        // Save the document to disk.
        doc->Save(outputPath, ooxmlSaveOptions);
        // ExEnd:UpdateLastSavedTimeProperty
        std::cout << "Updated Last Saved Time Property successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void KeepLegacyControlChars(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:KeepLegacyControlChars
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

        System::SharedPtr<OoxmlSaveOptions> so = System::MakeObject<OoxmlSaveOptions>(SaveFormat::FlatOpc);
        so->set_KeepLegacyControlChars(true);

        System::String outputPath = outputDataDir + u"WorkingWithOoxml.KeepLegacyControlChars.docx";
        // Save the document to disk.
        doc->Save(outputPath, so);

        // ExEnd:KeepLegacyControlChars
        std::cout << "Updated Last Saved With Keeping Legacy Control Chars Successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

	void SetCompressionLevel(System::String const& inputDataDir, System::String const& outputDataDir)
	{
		// ExStart:SetCompressionLevel
		System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

		System::SharedPtr<OoxmlSaveOptions> so = System::MakeObject<OoxmlSaveOptions>(SaveFormat::Docx);
		so->set_CompressionLevel(CompressionLevel::SuperFast);


		// Save the document to disk.
		System::String outputPath = outputDataDir + u"WorkingWithOoxml.SetCompressionLevel.docx";
		// Save the document to disk.
		doc->Save(outputPath, so);

		// ExEnd:SetCompressionLevel
		std::cout << "Document save with a Compression Level Successfully.\nFile saved at " << outputPath.ToUtf8String() << '\n';
	}
}

void WorkingWithOoxml()
{
    std::cout << "WorkingWithOoxml example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();
    EncryptDocxWithPassword(inputDataDir, outputDataDir);
    SetOOXMLCompliance(inputDataDir, outputDataDir);
    UpdateLastSavedTimeProperty(inputDataDir, outputDataDir);
    KeepLegacyControlChars(inputDataDir, outputDataDir);
    SetCompressionLevel(inputDataDir, outputDataDir);
    std::cout << "WorkingWithOoxml example finished." << std::endl << std::endl;
}