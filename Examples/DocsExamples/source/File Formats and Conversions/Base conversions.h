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
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <drawing/image.h>
#include <drawing/imaging/frame_dimension.h>
#include <system/array.h>
#include <system/details/dispose_guard.h>
#include <system/guid.h>
#include <system/io/file.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/io/stream.h>

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
        //ExStart:OpenDocument
        auto doc = MakeObject<Document>(MyDir + u"Document.doc");
        //ExEnd:OpenDocument

        doc->Save(ArtifactsDir + u"BaseConversions.DocToDocx.docx");
        //ExEnd:LoadAndSave
    }

    void DocxToRtf()
    {
        //ExStart:LoadAndSaveToStream
        //ExStart:OpeningFromStream
        // Read only access is enough for Aspose.Words to load a document.
        SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(MyDir + u"Document.docx");

        auto doc = MakeObject<Document>(stream);
        // You can close the stream now, it is no longer needed because the document is in memory.
        stream->Close();
        //ExEnd:OpeningFromStream

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
        //ExStart:SaveToMarkdownDocument
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Some text!");

        doc->Save(ArtifactsDir + u"BaseConversions.DocxToMarkdown.md");
        //ExEnd:SaveToMarkdownDocument
    }

    void DocxToTxt()
    {
        //ExStart:DocxToTxt
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        doc->Save(ArtifactsDir + u"BaseConversions.DocxToTxt.txt");
        //ExEnd:DocxToTxt
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
