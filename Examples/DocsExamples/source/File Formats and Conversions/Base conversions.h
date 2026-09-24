#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/ConvertUtil.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Replacing/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Saving/CompressionLevel.h>
#include <Aspose.Words.Cpp/Saving/XlsxSaveOptions.h>
#include <drawing/image.h>
#include <drawing/imaging/frame_dimension.h>
#include <system/array.h>
#include <system/details/dispose_guard.h>
#include <system/guid.h>
#include <system/io/file.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/io/stream.h>

#ifdef ASPOSE_EMAIL_AVAILABLE
// Aspose.Email for C++ headers. Its NuGet package puts include\Aspose.Email.Cpp
// on the include path, so these are referenced without a package prefix.
#include <Clients/Smtp/SmtpClient/SmtpClient.h>
#include <MailAddressCollection.h>
#include <MailMessage.h>
#include <MhtmlLoadOptions.h>
#endif

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace DocsExamples { namespace File_Formats_and_Conversions {

class BaseConversions : public DocsExamplesBase
{
public:
    void DocToDocx()
    {
        //ExStart:LoadAndSave
        //GistId:5938495d10e6b402f1e42ce5a45926ca
        //ExStart:OpenDocument
        auto doc = MakeObject<Document>(MyDir + u"Document.doc");
        //ExEnd:OpenDocument

        doc->Save(ArtifactsDir + u"BaseConversions.DocToDocx.docx");
        //ExEnd:LoadAndSave
    }

    void DocxToRtf()
    {
        //ExStart:LoadAndSaveToStream
        //GistId:5938495d10e6b402f1e42ce5a45926ca
        //ExStart:OpenFromStream
        //GistId:9ed4780658a7b003eda4e472d9270e3b
        // Read only access is enough for Aspose.Words to load a document.
        SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(MyDir + u"Document.docx");

        auto doc = MakeObject<Document>(stream);
        // You can close the stream now, it is no longer needed because the document is in memory.
        stream->Close();
        //ExEnd:OpenFromStream

        // ... do something with the document.

        // Convert the document to a different format and save to stream.
        auto dstStream = MakeObject<System::IO::MemoryStream>();
        doc->Save(dstStream, SaveFormat::Rtf);

        // Rewind the stream position back to zero so it is ready for the next reader.
        dstStream->set_Position(0);
        //ExEnd:LoadAndSaveToStream

        System::IO::File::WriteAllBytes(ArtifactsDir + u"BaseConversions.DocxToRtf.rtf", dstStream->ToArray());
    }

    void DocxToPdf()
    {
        //ExStart:DocxToPdf
        //GistId:b9784b73e288805e08fba6e3fc5ae2af
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        doc->Save(ArtifactsDir + u"BaseConversions.DocxToPdf.pdf");
        //ExEnd:DocxToPdf
    }

    void DocxToByte()
    {
        //ExStart:DocxToByte
        //GistId:38c10fec10086101e61bed049e52363c
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto outStream = MakeObject<System::IO::MemoryStream>();
        doc->Save(outStream, SaveFormat::Docx);

        ArrayPtr<uint8_t> docBytes = outStream->ToArray();
        auto inStream = MakeObject<System::IO::MemoryStream>(docBytes);

        auto docFromBytes = MakeObject<Document>(inStream);
        //ExEnd:DocxToByte
    }

    void DocxToEpub()
    {
        //ExStart:DocxToEpub
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        doc->Save(ArtifactsDir + u"BaseConversions.DocxToEpub.epub");
        //ExEnd:DocxToEpub
    }

    void DocxToMarkdown()
    {
        //ExStart:DocxToMarkdown
        //GistId:1739a7dc53ee2cce1ac97f4ef9fa7310
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Some text!");

        doc->Save(ArtifactsDir + u"BaseConversions.DocxToMarkdown.md");
        //ExEnd:DocxToMarkdown
    }

    void DocxToTxt()
    {
        //ExStart:DocxToTxt
        //GistId:922a9c5d9606a0c5cf0682b4aadfaf29
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        doc->Save(ArtifactsDir + u"BaseConversions.DocxToTxt.txt");
        //ExEnd:DocxToTxt
    }

    void DocxToHtml()
    {
        //ExStart:DocxToHtml
        //GistId:058a709735430153ce719761043ded8f
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        doc->Save(ArtifactsDir + u"BaseConversions.DocxToHtml.html");
        //ExEnd:DocxToHtml
    }

#ifdef ASPOSE_EMAIL_AVAILABLE
    void DocxToMhtml()
    {
        //ExStart:DocxToMhtml
        //GistId:ce77d307265430a38e77b141bb06f775
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<System::IO::Stream> stream = MakeObject<System::IO::MemoryStream>();
        doc->Save(stream, SaveFormat::Mhtml);

        // Rewind the stream to the beginning so Aspose.Email can read it.
        stream->set_Position(0);

        // Create an Aspose.Email MIME email message from the stream.
        SharedPtr<Aspose::Email::MailMessage> message =
            Aspose::Email::MailMessage::Load(stream, MakeObject<Aspose::Email::MhtmlLoadOptions>());
        message->set_From(u"your_from@email.com");
        message->get_To()->Add(u"your_to@email.com");
        message->set_Subject(u"Aspose.Words + Aspose.Email MHTML Test Message");

        // Send the message using Aspose.Email.
        auto client = MakeObject<Aspose::Email::Clients::Smtp::SmtpClient>();
        client->set_Host(u"your_smtp.com");
        client->Send(message);
        //ExEnd:DocxToMhtml
    }
#endif

    void DocxToXlsx()
    {
        //ExStart:DocxToXlsx
        //GistId:0833496410d6beb2913f35494196888e
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        doc->Save(ArtifactsDir + u"BaseConversions.DocxToXlsx.xlsx");
        //ExEnd:DocxToXlsx
    }

    void FindReplaceXlsx()
    {
        //ExStart:FindReplaceXlsx
        //GistId:b08318e6eb24ad14897dfce529b9b9fc
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Ruby bought a ruby necklace.");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        auto options = MakeObject<Replacing::FindReplaceOptions>();

        // Set the "MatchCase" flag to "true" to apply case sensitivity while finding strings to replace.
        // Set the "MatchCase" flag to "false" to ignore character case while searching for text to replace.
        options->set_MatchCase(true);

        doc->get_Range()->Replace(u"Ruby", u"Jade", options);

        doc->Save(ArtifactsDir + u"BaseConversions.FindReplaceXlsx.xlsx");
        //ExEnd:FindReplaceXlsx
    }

    void CompressXlsx()
    {
        //ExStart:CompressXlsx
        //GistId:b08318e6eb24ad14897dfce529b9b9fc
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto saveOptions = MakeObject<Saving::XlsxSaveOptions>();
        saveOptions->set_CompressionLevel(Saving::CompressionLevel::Maximum);

        doc->Save(ArtifactsDir + u"BaseConversions.CompressXlsx.xlsx", saveOptions);
        //ExEnd:CompressXlsx
    }

    void TxtToDocx()
    {
        //ExStart:TxtToDocx
        // The encoding of the text file is automatically detected.
        auto doc = MakeObject<Document>(MyDir + u"English text.txt");
        doc->Save(ArtifactsDir + u"BaseConversions.TxtToDocx.docx");
        //ExEnd:TxtToDocx
    }

    void ImagesToPdf()
    {
        //ExStart:ImageToPdf
        //GistId:b9784b73e288805e08fba6e3fc5ae2af
        ConvertImageToPdf(ImagesDir + u"Logo.jpg", ArtifactsDir + u"BaseConversions.JpgToPdf.pdf");
        ConvertImageToPdf(ImagesDir + u"Transparent background logo.png", ArtifactsDir + u"BaseConversions.PngToPdf.pdf");
        ConvertImageToPdf(ImagesDir + u"Windows MetaFile.wmf", ArtifactsDir + u"BaseConversions.WmfToPdf.pdf");
        ConvertImageToPdf(ImagesDir + u"Tagged Image File Format.tiff", ArtifactsDir + u"BaseConversions.TiffToPdf.pdf");
        ConvertImageToPdf(ImagesDir + u"Graphics Interchange Format.gif", ArtifactsDir + u"BaseConversions.GifToPdf.pdf");
        //ExEnd:ImageToPdf
    }

    //ExStart:ConvertImageToPdf
    //GistId:b9784b73e288805e08fba6e3fc5ae2af
    /// <summary>
    /// Converts an image to PDF using Aspose.Words for .NET.
    /// </summary>
    /// <param name="inputFileName">File name of input image file.</param>
    /// <param name="outputFileName">Output PDF file name.</param>
    void ConvertImageToPdf(String inputFileName, String outputFileName)
    {
        std::cout << (String(u"Converting ") + inputFileName + u" to PDF ....") << std::endl;
                
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Read the image from file, ensure it is disposed.
        {
            SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromFile(inputFileName);

			// Insert a section break before each new page, in case of a multi-frame TIFF.
			builder->InsertBreak(BreakType::SectionBreakNewPage);

			// We want the size of the page to be the same as the size of the image.
			// Convert pixels to points to size the page to the actual image size.
			SharedPtr<PageSetup> ps = builder->get_PageSetup();
			ps->set_PageWidth(ConvertUtil::PixelToPoint(image->get_Width(), image->get_HorizontalResolution()));
			ps->set_PageHeight(ConvertUtil::PixelToPoint(image->get_Height(), image->get_VerticalResolution()));

			// Insert the image into the document and position it at the top left corner of the page.
			builder->InsertImage(image, RelativeHorizontalPosition::Page, 0, RelativeVerticalPosition::Page, 0, ps->get_PageWidth(), ps->get_PageHeight(), WrapType::None);
        }

        doc->Save(outputFileName);        
    }
    //ExEnd:ConvertImageToPdf
};

}} // namespace DocsExamples::File_Formats_and_Conversions
