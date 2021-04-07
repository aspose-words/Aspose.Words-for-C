#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/CertificateHolder.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/FileCorruptedException.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/LoadFormat.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/file_stream.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/encoding.h>
#include <system/text/regularexpressions/regex.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExFile : public ApiExampleBase
{
public:
    void CatchFileCorruptedException()
    {
        //ExStart
        //ExFor:FileCorruptedException
        //ExSummary:Shows how to catch a FileCorruptedException.
        try
        {
            // If we get an "Unreadable content" error message when trying to open a document using Microsoft Word,
            // chances are that we will get an exception thrown when trying to load that document using Aspose.Words.
            auto doc = MakeObject<Document>(MyDir + u"Corrupted document.docx");
        }
        catch (FileCorruptedException& e)
        {
            std::cout << e->get_Message() << std::endl;
        }

        //ExEnd
    }

    void DetectEncoding()
    {
        //ExStart
        //ExFor:FileFormatInfo.Encoding
        //ExFor:FileFormatUtil
        //ExSummary:Shows how to detect encoding in an html file.
        SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(MyDir + u"Document.html");

        ASSERT_EQ(LoadFormat::Html, info->get_LoadFormat());

        // The Encoding property is used only when we create a FileFormatInfo object for an html document.
        ASSERT_EQ(u"Windows-1252", info->get_Encoding()->get_WebName());
        ASSERT_EQ(1252, info->get_Encoding()->get_CodePage());
        //ExEnd

        info = FileFormatUtil::DetectFileFormat(MyDir + u"Document.docx");

        ASSERT_EQ(LoadFormat::Docx, info->get_LoadFormat());
        ASSERT_TRUE(info->get_Encoding() == nullptr);
    }

    void FileFormatToString()
    {
        //ExStart
        //ExFor:FileFormatUtil.ContentTypeToLoadFormat(String)
        //ExFor:FileFormatUtil.ContentTypeToSaveFormat(String)
        //ExSummary:Shows how to find the corresponding Aspose load/save format from each media type string.
        // The ContentTypeToSaveFormat/ContentTypeToLoadFormat methods only accept official IANA media type names, also known as MIME types.
        // All valid media types are listed here: https://www.iana.org/assignments/media-types/media-types.xhtml.

        // Trying to associate a SaveFormat with a partial media type string will not work.
        ASSERT_THROW(FileFormatUtil::ContentTypeToSaveFormat(u"jpeg"), System::ArgumentException);

        // If Aspose.Words does not have a corresponding save/load format for a content type, an exception will also be thrown.
        ASSERT_THROW(FileFormatUtil::ContentTypeToSaveFormat(u"application/zip"), System::ArgumentException);

        // Files of the types listed below can be saved, but not loaded using Aspose.Words.
        ASSERT_THROW(FileFormatUtil::ContentTypeToLoadFormat(u"image/jpeg"), System::ArgumentException);

        ASSERT_EQ(SaveFormat::Jpeg, FileFormatUtil::ContentTypeToSaveFormat(u"image/jpeg"));
        ASSERT_EQ(SaveFormat::Png, FileFormatUtil::ContentTypeToSaveFormat(u"image/png"));
        ASSERT_EQ(SaveFormat::Tiff, FileFormatUtil::ContentTypeToSaveFormat(u"image/tiff"));
        ASSERT_EQ(SaveFormat::Gif, FileFormatUtil::ContentTypeToSaveFormat(u"image/gif"));
        ASSERT_EQ(SaveFormat::Emf, FileFormatUtil::ContentTypeToSaveFormat(u"image/x-emf"));
        ASSERT_EQ(SaveFormat::Xps, FileFormatUtil::ContentTypeToSaveFormat(u"application/vnd.ms-xpsdocument"));
        ASSERT_EQ(SaveFormat::Pdf, FileFormatUtil::ContentTypeToSaveFormat(u"application/pdf"));
        ASSERT_EQ(SaveFormat::Svg, FileFormatUtil::ContentTypeToSaveFormat(u"image/svg+xml"));
        ASSERT_EQ(SaveFormat::Epub, FileFormatUtil::ContentTypeToSaveFormat(u"application/epub+zip"));

        // For file types that can be saved and loaded, we can match a media type to both a load format and a save format.
        ASSERT_EQ(LoadFormat::Doc, FileFormatUtil::ContentTypeToLoadFormat(u"application/msword"));
        ASSERT_EQ(SaveFormat::Doc, FileFormatUtil::ContentTypeToSaveFormat(u"application/msword"));

        ASSERT_EQ(LoadFormat::Docx, FileFormatUtil::ContentTypeToLoadFormat(u"application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
        ASSERT_EQ(SaveFormat::Docx, FileFormatUtil::ContentTypeToSaveFormat(u"application/vnd.openxmlformats-officedocument.wordprocessingml.document"));

        ASSERT_EQ(LoadFormat::Text, FileFormatUtil::ContentTypeToLoadFormat(u"text/plain"));
        ASSERT_EQ(SaveFormat::Text, FileFormatUtil::ContentTypeToSaveFormat(u"text/plain"));

        ASSERT_EQ(LoadFormat::Rtf, FileFormatUtil::ContentTypeToLoadFormat(u"application/rtf"));
        ASSERT_EQ(SaveFormat::Rtf, FileFormatUtil::ContentTypeToSaveFormat(u"application/rtf"));

        ASSERT_EQ(LoadFormat::Html, FileFormatUtil::ContentTypeToLoadFormat(u"text/html"));
        ASSERT_EQ(SaveFormat::Html, FileFormatUtil::ContentTypeToSaveFormat(u"text/html"));

        ASSERT_EQ(LoadFormat::Mhtml, FileFormatUtil::ContentTypeToLoadFormat(u"multipart/related"));
        ASSERT_EQ(SaveFormat::Mhtml, FileFormatUtil::ContentTypeToSaveFormat(u"multipart/related"));
        //ExEnd
    }

    void DetectDocumentEncryption()
    {
        //ExStart
        //ExFor:FileFormatUtil.DetectFileFormat(String)
        //ExFor:FileFormatInfo
        //ExFor:FileFormatInfo.LoadFormat
        //ExFor:FileFormatInfo.IsEncrypted
        //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and encryption.
        auto doc = MakeObject<Document>();

        // Configure a SaveOptions object to encrypt the document
        // with a password when we save it, and then save the document.
        auto saveOptions = MakeObject<OdtSaveOptions>(SaveFormat::Odt);
        saveOptions->set_Password(u"MyPassword");

        doc->Save(ArtifactsDir + u"File.DetectDocumentEncryption.odt", saveOptions);

        // Verify the file type of our document, and its encryption status.
        SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(ArtifactsDir + u"File.DetectDocumentEncryption.odt");

        ASSERT_EQ(u".odt", FileFormatUtil::LoadFormatToExtension(info->get_LoadFormat()));
        ASSERT_TRUE(info->get_IsEncrypted());
        //ExEnd
    }

    void DetectDigitalSignatures()
    {
        //ExStart
        //ExFor:FileFormatUtil.DetectFileFormat(String)
        //ExFor:FileFormatInfo
        //ExFor:FileFormatInfo.LoadFormat
        //ExFor:FileFormatInfo.HasDigitalSignature
        //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and presence of digital signatures.
        // Use a FileFormatInfo instance to verify that a document is not digitally signed.
        SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(MyDir + u"Document.docx");

        ASSERT_EQ(u".docx", FileFormatUtil::LoadFormatToExtension(info->get_LoadFormat()));
        ASSERT_FALSE(info->get_HasDigitalSignature());

        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw", nullptr);
        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_SignTime(System::DateTime::get_Now());
        DigitalSignatureUtil::Sign(MyDir + u"Document.docx", ArtifactsDir + u"File.DetectDigitalSignatures.docx", certificateHolder, signOptions);

        // Use a new FileFormatInstance to confirm that it is signed.
        info = FileFormatUtil::DetectFileFormat(ArtifactsDir + u"File.DetectDigitalSignatures.docx");

        ASSERT_TRUE(info->get_HasDigitalSignature());

        // We can load and access the signatures of a signed document in a collection like this.
        ASSERT_EQ(1, DigitalSignatureUtil::LoadSignatures(ArtifactsDir + u"File.DetectDigitalSignatures.docx")->get_Count());
        //ExEnd
    }

    void SaveToDetectedFileFormat()
    {
        //ExStart
        //ExFor:FileFormatUtil.DetectFileFormat(Stream)
        //ExFor:FileFormatUtil.LoadFormatToExtension(LoadFormat)
        //ExFor:FileFormatUtil.ExtensionToSaveFormat(String)
        //ExFor:FileFormatUtil.SaveFormatToExtension(SaveFormat)
        //ExFor:FileFormatUtil.LoadFormatToSaveFormat(LoadFormat)
        //ExFor:Document.OriginalFileName
        //ExFor:FileFormatInfo.LoadFormat
        //ExFor:LoadFormat
        //ExSummary:Shows how to use the FileFormatUtil methods to detect the format of a document.
        // Load a document from a file that is missing a file extension, and then detect its file format.
        {
            SharedPtr<System::IO::FileStream> docStream = System::IO::File::OpenRead(MyDir + u"Word document with missing file extension");
            SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(docStream);
            LoadFormat loadFormat = info->get_LoadFormat();

            ASSERT_EQ(LoadFormat::Doc, loadFormat);

            // Below are two methods of converting a LoadFormat to its corresponding SaveFormat.
            // 1 -  Get the file extension string for the LoadFormat, then get the corresponding SaveFormat from that string:
            String fileExtension = FileFormatUtil::LoadFormatToExtension(loadFormat);
            SaveFormat saveFormat = FileFormatUtil::ExtensionToSaveFormat(fileExtension);

            // 2 -  Convert the LoadFormat directly to its SaveFormat:
            saveFormat = FileFormatUtil::LoadFormatToSaveFormat(loadFormat);

            // Load a document from the stream, and then save it to the automatically detected file extension.
            auto doc = MakeObject<Document>(docStream);

            ASSERT_EQ(u".doc", FileFormatUtil::SaveFormatToExtension(saveFormat));

            doc->Save(ArtifactsDir + u"File.SaveToDetectedFileFormat" + FileFormatUtil::SaveFormatToExtension(saveFormat));
        }
        //ExEnd
    }

    void DetectFileFormat_SaveFormatToLoadFormat()
    {
        //ExStart
        //ExFor:FileFormatUtil.SaveFormatToLoadFormat(SaveFormat)
        //ExSummary:Shows how to convert a save format to its corresponding load format.
        ASSERT_EQ(LoadFormat::Html, FileFormatUtil::SaveFormatToLoadFormat(SaveFormat::Html));

        // Some file types can have documents saved to, but not loaded from using Aspose.Words.
        // If we attempt to convert a save format of such a type to a load format, an exception will be thrown.
        ASSERT_THROW(FileFormatUtil::SaveFormatToLoadFormat(SaveFormat::Jpeg), System::ArgumentException);
        //ExEnd
    }

    void ExtractImages()
    {
        //ExStart
        //ExFor:Shape
        //ExFor:Shape.ImageData
        //ExFor:Shape.HasImage
        //ExFor:ImageData
        //ExFor:FileFormatUtil.ImageTypeToExtension(ImageType)
        //ExFor:ImageData.ImageType
        //ExFor:ImageData.Save(String)
        //ExFor:CompositeNode.GetChildNodes(NodeType, bool)
        //ExSummary:Shows how to extract images from a document, and save them to the local file system as individual files.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        // Get the collection of shapes from the document,
        // and save the image data of every shape with an image as a file to the local file system.
        SharedPtr<NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);

        ASSERT_EQ(9, shapes->LINQ_Count([](SharedPtr<Node> s) { return (System::DynamicCast<Shape>(s))->get_HasImage(); }));

        int imageIndex = 0;
        for (auto shape : System::IterateOver(shapes->LINQ_OfType<SharedPtr<Shape>>()))
        {
            if (shape->get_HasImage())
            {
                // The image data of shapes may contain images of many possible image formats.
                // We can determine a file extension for each image automatically, based on its format.
                String imageFileName =
                    String::Format(u"File.ExtractImages.{0}{1}", imageIndex, FileFormatUtil::ImageTypeToExtension(shape->get_ImageData()->get_ImageType()));
                shape->get_ImageData()->Save(ArtifactsDir + imageFileName);
                imageIndex++;
            }
        }
        //ExEnd

        ASSERT_EQ(9, System::IO::Directory::GetFiles(ArtifactsDir)
                         ->LINQ_Count(static_cast<System::Func<String, bool>>(static_cast<std::function<bool(String s)>>(
                             [](String s) -> bool {
                                 return System::Text::RegularExpressions::Regex::IsMatch(s, u"^.+\\.(jpeg|png|emf|wmf)$") &&
                                        s.StartsWith(ArtifactsDir + u"File.ExtractImages");
                             }))));
    }
};

} // namespace ApiExamples
