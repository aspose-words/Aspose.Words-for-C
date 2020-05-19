#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <MailMessage.h>
#include <MailAddressCollection.h>
//#include <cstdint>
#include <Clients/Smtp/SmtpClient/SmtpClient.h>
//#include <Clients/SecurityOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;


namespace
{
	void ConvertDocxToHtml(System::String const& inputDataDir, System::String const& outputDataDir)
	{
		// ExStart:ConvertDocxToHtml
		// Load the document from disk.
		System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Test File (docx).docx");

		// Save the document into HTML.
		doc->Save(outputDataDir + u"Document_out.html", SaveFormat::Html);
		// ExEnd:ConvertDocxToHtml
		std::cout << "Document converted to html successfully." << std::endl << std::endl;
	}

    void ConvertDocumentToHtmlWithRoundtrip(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:ConvertDocumentToHtmlWithRoundtrip
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Test File (doc).doc");

        System::SharedPtr<HtmlSaveOptions> options = System::MakeObject<HtmlSaveOptions>();
        // HtmlSaveOptions.ExportRoundtripInformation property specifies
        // Whether to write the roundtrip information when saving to HTML, MHTML or EPUB.
        // Default value is true for HTML and false for MHTML and EPUB.
        options->set_ExportRoundtripInformation(true);

        doc->Save(outputDataDir + u"ConvertDocumentToHtmlWithRoundtrip.html", options);
        // ExEnd:ConvertDocumentToHtmlWithRoundtrip
        std::cout << "Document converted to html with roundtrip informations successfully." << std::endl;
        std::cout << "ConvertDocumentToHtmlWithRoundtrip example finished." << std::endl << std::endl;
    }

    void ExportResourcesUsingHtmlSaveOptions(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:ExportResourcesUsingHtmlSaveOptions
        // The path to the documents directory.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

        System::SharedPtr<HtmlSaveOptions> saveOptions = System::MakeObject<HtmlSaveOptions>();
        saveOptions->set_CssStyleSheetType(CssStyleSheetType::External);
        saveOptions->set_ExportFontResources(true);
        saveOptions->set_ResourceFolder(outputDataDir + u"\\Resources");
        
        doc->Save(outputDataDir + u"ExportResourcesUsingHtmlSaveOptions.html", saveOptions);
        // ExEnd:ExportResourcesUsingHtmlSaveOptions
        std::cout << "Save option specified successfully." << std::endl << "File saved at " << (outputDataDir + u"ExportResourcesUsingHtmlSaveOptions.html").ToUtf8String() << std::endl;
        std::cout << "ExportResourcesUsingHtmlSaveOptions example finished." << std::endl << std::endl;
    }

    void ExportFontsAsBase64(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:ExportFontsAsBase64
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

        System::SharedPtr<HtmlSaveOptions> saveOptions = System::MakeObject<HtmlSaveOptions>();
        saveOptions->set_ExportFontResources(true);
        saveOptions->set_ExportFontsAsBase64(true);
        
        System::String outputPath = outputDataDir + u"ExportFontsAsBase64.html";
        doc->Save(outputPath, saveOptions);
        // ExEnd:ExportFontsAsBase64
        std::cout << "Save option specified successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
        std::cout << "ExportFontsAsBase64 example finished." << std::endl << std::endl;
    }

    void ConvertDocumentToMhtmlAndEmail(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart: ConvertDocumentToMhtmlAndEmail
        // For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-C
        // Load the document into Aspose.Words.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Test File (docx).docx");

        // Save into a memory stream in MHTML format.
        System::SharedPtr<System::IO::MemoryStream> stream = System::MakeObject<System::IO::MemoryStream>();
        doc->Save(stream, SaveFormat::Mhtml);

        // Rewind the stream to the beginning so Aspose.Email can read it.
        stream->set_Position(0);

        // Create an Aspose.Network MIME email message from the stream.
        System::SharedPtr<Aspose::Email::MailMessage > message = System::MakeObject<Aspose::Email::MailMessage>();
        //message->Load(stream, System::MakeObject<Aspose::Email::MhtmlLoadOptions>());

        message->set_From(u"sender@sender.com");
        message->get_To()->Add(u"receiver@gmail.com");
        message->set_Subject(u"Aspose.Words + Aspose.Email MHTML Test Message");

        // Send the message using Aspose.Email
        System::SharedPtr<Aspose::Email::Clients::Smtp::SmtpClient> client = System::MakeObject<Aspose::Email::Clients::Smtp::SmtpClient>();
        client->set_Host(u"mail.server.com");
        client->set_Username(u"username");
        client->set_Password(u"password");
        client->set_Port(587);
        client->Send(message);
        // ExEnd: ConvertDocumentToMhtmlAndEmail
        std::cout << "ConvertDocumentToMhtmlAndEmail example finished." << std::endl << std::endl;
    }
}

void ConvertDocumentToHTML()
{
    std::cout << "ConvertDocumentToEPUB example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

	ConvertDocxToHtml(inputDataDir, outputDataDir);
    ConvertDocumentToHtmlWithRoundtrip(inputDataDir, outputDataDir);
    ExportResourcesUsingHtmlSaveOptions(inputDataDir, outputDataDir);
    ExportFontsAsBase64(inputDataDir, outputDataDir);
    ConvertDocumentToMhtmlAndEmail(inputDataDir, outputDataDir);
    std::cout << "ConvertDocumentToEPUB example finished." << std::endl << std::endl;
}