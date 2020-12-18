#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/CertificateHolder.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/FileCorruptedException.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/LoadFormat.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/SignOptions.h>
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

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
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
        // 'DetectFileFormat' not working on a non-html files
        SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(MyDir + u"Document.docx");
        ASSERT_EQ(LoadFormat::Docx, info->get_LoadFormat());
        ASSERT_TRUE(info->get_Encoding() == nullptr);

        // This time the property will not be null
        info = FileFormatUtil::DetectFileFormat(MyDir + u"Document.html");
        ASSERT_EQ(LoadFormat::Html, info->get_LoadFormat());
        ASSERT_FALSE(info->get_Encoding() == nullptr);
        //ExEnd
    }

    void FileFormatToString()
    {
        //ExStart
        //ExFor:FileFormatUtil.ContentTypeToLoadFormat(String)
        //ExFor:FileFormatUtil.ContentTypeToSaveFormat(String)
        //ExSummary:Shows how to find the corresponding Aspose load/save format from an IANA content type string.
        // Trying to search for a SaveFormat with a simple string will not work
        try
        {
            ASSERT_EQ(SaveFormat::Jpeg, FileFormatUtil::ContentTypeToSaveFormat(u"jpeg"));
        }
        catch (System::ArgumentException& e)
        {
            std::cout << e->get_Message() << std::endl;
        }

        // The convertion methods only accept official IANA type names, which are all listed here:
        //      https://www.iana.org/assignments/media-types/media-types.xhtml
        // Note that if a corresponding SaveFormat or LoadFormat for a type from that list does not exist in the Aspose enums,
        // converting will raise an exception just like in the code above

        // File types that can be saved to but not opened as documents will not have corresponding load formats
        // Attempting to convert them to load formats will raise an exception
        ASSERT_EQ(SaveFormat::Jpeg, FileFormatUtil::ContentTypeToSaveFormat(u"image/jpeg"));
        ASSERT_EQ(SaveFormat::Png, FileFormatUtil::ContentTypeToSaveFormat(u"image/png"));
        ASSERT_EQ(SaveFormat::Tiff, FileFormatUtil::ContentTypeToSaveFormat(u"image/tiff"));
        ASSERT_EQ(SaveFormat::Gif, FileFormatUtil::ContentTypeToSaveFormat(u"image/gif"));
        ASSERT_EQ(SaveFormat::Emf, FileFormatUtil::ContentTypeToSaveFormat(u"image/x-emf"));
        ASSERT_EQ(SaveFormat::Xps, FileFormatUtil::ContentTypeToSaveFormat(u"application/vnd.ms-xpsdocument"));
        ASSERT_EQ(SaveFormat::Pdf, FileFormatUtil::ContentTypeToSaveFormat(u"application/pdf"));
        ASSERT_EQ(SaveFormat::Svg, FileFormatUtil::ContentTypeToSaveFormat(u"image/svg+xml"));
        ASSERT_EQ(SaveFormat::Epub, FileFormatUtil::ContentTypeToSaveFormat(u"application/epub+zip"));

        // File types that can both be loaded and saved have corresponding load and save formats
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

        // Save it as an encrypted .odt
        auto saveOptions = MakeObject<OdtSaveOptions>(SaveFormat::Odt);
        saveOptions->set_Password(u"MyPassword");

        doc->Save(ArtifactsDir + u"File.DetectDocumentEncryption.odt", saveOptions);

        // Create a FileFormatInfo object for this document
        SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(ArtifactsDir + u"File.DetectDocumentEncryption.odt");

        // Verify the file type of our document and its encryption status
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
        // Use a FileFormatInfo instance to verify that a document is not digitally signed
        SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(MyDir + u"Document.docx");

        ASSERT_EQ(u".docx", FileFormatUtil::LoadFormatToExtension(info->get_LoadFormat()));
        ASSERT_FALSE(info->get_HasDigitalSignature());

        // Sign the document
        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw", nullptr);
        DigitalSignatureUtil::Sign(MyDir + u"Document.docx", ArtifactsDir + u"File.DetectDigitalSignatures.docx", certificateHolder,
                                   [&]
                                   {
                                       auto tmp_0 = MakeObject<SignOptions>();
                                       tmp_0->set_SignTime(System::DateTime::get_Now());
                                       return tmp_0;
                                   }());

        // Use a new FileFormatInstance to confirm that it is signed
        info = FileFormatUtil::DetectFileFormat(ArtifactsDir + u"File.DetectDigitalSignatures.docx");

        ASSERT_TRUE(info->get_HasDigitalSignature());

        // The signatures can then be accessed like this
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
        //ExSummary:Shows how to use the FileFormatUtil methods to detect the format of a document without any extension and save it with the correct file extension.
        // Load the document without a file extension into a stream and use the DetectFileFormat method to detect it's format
        // These are both times where you might need extract the file format as it's not visible
        // The file format of this document is actually ".doc"
        SharedPtr<System::IO::FileStream> docStream = System::IO::File::OpenRead(MyDir + u"Word document with missing file extension");
        SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(docStream);

        // Retrieve the LoadFormat of the document
        LoadFormat loadFormat = info->get_LoadFormat();

        // There are two methods of converting LoadFormat enumerations to SaveFormat enumerations
        // Method #1
        // Convert the LoadFormat to a String first for working with. The String will include the leading dot in front of the extension
        String fileExtension = FileFormatUtil::LoadFormatToExtension(loadFormat);
        // Now convert this extension into the corresponding SaveFormat enumeration
        SaveFormat saveFormat = FileFormatUtil::ExtensionToSaveFormat(fileExtension);

        // Method #2
        // Convert the LoadFormat enumeration directly to the SaveFormat enumeration
        saveFormat = FileFormatUtil::LoadFormatToSaveFormat(loadFormat);

        // Load a document from the stream.
        auto doc = MakeObject<Document>(docStream);

        // Save the document with the original file name, " Out" and the document's file extension
        doc->Save(ArtifactsDir + u"File.SaveToDetectedFileFormat" + FileFormatUtil::SaveFormatToExtension(saveFormat));
        //ExEnd

        ASSERT_EQ(u".doc", FileFormatUtil::SaveFormatToExtension(saveFormat));
    }

    void DetectFileFormat_SaveFormatToLoadFormat()
    {
        //ExStart
        //ExFor:FileFormatUtil.SaveFormatToLoadFormat(SaveFormat)
        //ExSummary:Shows how to use the FileFormatUtil class and to convert a SaveFormat enumeration into the corresponding LoadFormat enumeration.
        // Define the SaveFormat enumeration to convert
        const SaveFormat saveFormat = SaveFormat::Html;
        // Convert the SaveFormat enumeration to LoadFormat enumeration
        LoadFormat loadFormat = FileFormatUtil::SaveFormatToLoadFormat(saveFormat);
        std::cout << (String(u"The converted LoadFormat is: ") + FileFormatUtil::LoadFormatToExtension(loadFormat)) << std::endl;
        //ExEnd

        ASSERT_EQ(u".html", FileFormatUtil::SaveFormatToExtension(saveFormat));
        ASSERT_EQ(u".html", FileFormatUtil::LoadFormatToExtension(loadFormat));
    }

    void ExtractImagesToFiles()
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
        //ExSummary:Shows how to extract images from a document and save them as files.
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        SharedPtr<NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);

        ASSERT_EQ(9, shapes->LINQ_Count([](SharedPtr<Node> s) { return (System::DynamicCast<Shape>(s))->get_HasImage(); }));

        int imageIndex = 0;
        for (auto shape : System::IterateOver(shapes->LINQ_OfType<SharedPtr<Shape>>()))
        {
            if (shape->get_HasImage())
            {
                String imageFileName = String::Format(u"File.ExtractImagesToFiles.{0}{1}", imageIndex,
                                                                      FileFormatUtil::ImageTypeToExtension(shape->get_ImageData()->get_ImageType()));
                shape->get_ImageData()->Save(ArtifactsDir + imageFileName);
                imageIndex++;
            }
        }
        //ExEnd

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(String s)> _local_func_1 = [](String s)
        {
            return System::Text::RegularExpressions::Regex::IsMatch(s, u"^.+\\.(jpeg|png|emf|wmf)$") &&
                   s.StartsWith(ArtifactsDir + u"File.ExtractImagesToFiles");
        };

        ASSERT_EQ(9, System::IO::Directory::GetFiles(ArtifactsDir)->LINQ_Count(static_cast<System::Func<String, bool>>(_local_func_1)));
    }
};

} // namespace ApiExamples
