// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExFile.h"

#include <testing/test_predicates.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/io/file_stream.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/date_time.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/LoadFormat.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>
#include <Aspose.Words.Cpp/Model/Document/FileCorruptedException.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/CertificateHolder.h>


using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2600632204u, ::Aspose::Words::ApiExamples::ExFile, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExFile : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExFile> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExFile>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExFile> ExFile::s_instance;

} // namespace gtest_test

void ExFile::CatchFileCorruptedException()
{
    //ExStart
    //ExFor:FileCorruptedException
    //ExSummary:Shows how to catch a FileCorruptedException.
    try
    {
        // If we get an "Unreadable content" error message when trying to open a document using Microsoft Word,
        // chances are that we will get an exception thrown when trying to load that document using Aspose.Words.
        auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Corrupted document.docx");
    }
    catch (Aspose::Words::FileCorruptedException& e)
    {
        std::cout << e->get_Message() << std::endl;
    }
    
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFile, CatchFileCorruptedException)
{
    s_instance->CatchFileCorruptedException();
}

} // namespace gtest_test

void ExFile::DetectEncoding()
{
    //ExStart
    //ExFor:FileFormatInfo.Encoding
    //ExFor:FileFormatUtil
    //ExSummary:Shows how to detect encoding in an html file.
    System::SharedPtr<Aspose::Words::FileFormatInfo> info = Aspose::Words::FileFormatUtil::DetectFileFormat(get_MyDir() + u"Document.html");
    
    ASSERT_EQ(Aspose::Words::LoadFormat::Html, info->get_LoadFormat());
    
    // The Encoding property is used only when we create a FileFormatInfo object for an html document.
    ASSERT_EQ(u"Western European (Windows)", info->get_Encoding()->get_EncodingName());
    ASSERT_EQ(1252, info->get_Encoding()->get_CodePage());
    //ExEnd
    
    info = Aspose::Words::FileFormatUtil::DetectFileFormat(get_MyDir() + u"Document.docx");
    
    ASSERT_EQ(Aspose::Words::LoadFormat::Docx, info->get_LoadFormat());
    ASSERT_TRUE(System::TestTools::IsNull(info->get_Encoding()));
}

namespace gtest_test
{

TEST_F(ExFile, DetectEncoding)
{
    s_instance->DetectEncoding();
}

} // namespace gtest_test

void ExFile::FileFormatToString()
{
    //ExStart
    //ExFor:FileFormatUtil.ContentTypeToLoadFormat(String)
    //ExFor:FileFormatUtil.ContentTypeToSaveFormat(String)
    //ExSummary:Shows how to find the corresponding Aspose load/save format from each media type string.
    // The ContentTypeToSaveFormat/ContentTypeToLoadFormat methods only accept official IANA media type names, also known as MIME types. 
    // All valid media types are listed here: https://www.iana.org/assignments/media-types/media-types.xhtml.
    
    // Trying to associate a SaveFormat with a partial media type string will not work.
    ASSERT_THROW(static_cast<std::function<void()>>([]() -> void
    {
        Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"jpeg");
    })(), System::ArgumentException);
    
    // If Aspose.Words does not have a corresponding save/load format for a content type, an exception will also be thrown.
    ASSERT_THROW(static_cast<std::function<void()>>([]() -> void
    {
        Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"application/zip");
    })(), System::ArgumentException);
    
    // Files of the types listed below can be saved, but not loaded using Aspose.Words.
    ASSERT_THROW(static_cast<std::function<void()>>([]() -> void
    {
        Aspose::Words::FileFormatUtil::ContentTypeToLoadFormat(u"image/jpeg");
    })(), System::ArgumentException);
    
    ASSERT_EQ(Aspose::Words::SaveFormat::Jpeg, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"image/jpeg"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Png, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"image/png"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Tiff, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"image/tiff"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Gif, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"image/gif"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Emf, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"image/x-emf"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Xps, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"application/vnd.ms-xpsdocument"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Pdf, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"application/pdf"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Svg, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"image/svg+xml"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Epub, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"application/epub+zip"));
    
    // For file types that can be saved and loaded, we can match a media type to both a load format and a save format.
    ASSERT_EQ(Aspose::Words::LoadFormat::Doc, Aspose::Words::FileFormatUtil::ContentTypeToLoadFormat(u"application/msword"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Doc, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"application/msword"));
    
    ASSERT_EQ(Aspose::Words::LoadFormat::Docx, Aspose::Words::FileFormatUtil::ContentTypeToLoadFormat(u"application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Docx, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
    
    ASSERT_EQ(Aspose::Words::LoadFormat::Text, Aspose::Words::FileFormatUtil::ContentTypeToLoadFormat(u"text/plain"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Text, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"text/plain"));
    
    ASSERT_EQ(Aspose::Words::LoadFormat::Rtf, Aspose::Words::FileFormatUtil::ContentTypeToLoadFormat(u"application/rtf"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Rtf, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"application/rtf"));
    
    ASSERT_EQ(Aspose::Words::LoadFormat::Html, Aspose::Words::FileFormatUtil::ContentTypeToLoadFormat(u"text/html"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Html, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"text/html"));
    
    ASSERT_EQ(Aspose::Words::LoadFormat::Mhtml, Aspose::Words::FileFormatUtil::ContentTypeToLoadFormat(u"multipart/related"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Mhtml, Aspose::Words::FileFormatUtil::ContentTypeToSaveFormat(u"multipart/related"));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFile, FileFormatToString)
{
    s_instance->FileFormatToString();
}

} // namespace gtest_test

void ExFile::DetectDocumentEncryption()
{
    //ExStart
    //ExFor:FileFormatUtil.DetectFileFormat(String)
    //ExFor:FileFormatInfo
    //ExFor:FileFormatInfo.LoadFormat
    //ExFor:FileFormatInfo.IsEncrypted
    //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and encryption.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Configure a SaveOptions object to encrypt the document
    // with a password when we save it, and then save the document.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OdtSaveOptions>(Aspose::Words::SaveFormat::Odt);
    saveOptions->set_Password(u"MyPassword");
    
    doc->Save(get_ArtifactsDir() + u"File.DetectDocumentEncryption.odt", saveOptions);
    
    // Verify the file type of our document, and its encryption status.
    System::SharedPtr<Aspose::Words::FileFormatInfo> info = Aspose::Words::FileFormatUtil::DetectFileFormat(get_ArtifactsDir() + u"File.DetectDocumentEncryption.odt");
    
    ASSERT_EQ(u".odt", Aspose::Words::FileFormatUtil::LoadFormatToExtension(info->get_LoadFormat()));
    ASSERT_TRUE(info->get_IsEncrypted());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFile, DetectDocumentEncryption)
{
    s_instance->DetectDocumentEncryption();
}

} // namespace gtest_test

void ExFile::DetectDigitalSignatures()
{
    //ExStart
    //ExFor:FileFormatUtil.DetectFileFormat(String)
    //ExFor:FileFormatInfo
    //ExFor:FileFormatInfo.LoadFormat
    //ExFor:FileFormatInfo.HasDigitalSignature
    //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and presence of digital signatures.
    // Use a FileFormatInfo instance to verify that a document is not digitally signed.
    System::SharedPtr<Aspose::Words::FileFormatInfo> info = Aspose::Words::FileFormatUtil::DetectFileFormat(get_MyDir() + u"Document.docx");
    
    ASSERT_EQ(u".docx", Aspose::Words::FileFormatUtil::LoadFormatToExtension(info->get_LoadFormat()));
    ASSERT_FALSE(info->get_HasDigitalSignature());
    
    System::SharedPtr<Aspose::Words::DigitalSignatures::CertificateHolder> certificateHolder = Aspose::Words::DigitalSignatures::CertificateHolder::Create(get_MyDir() + u"morzal.pfx", u"aw", nullptr);
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_SignTime(System::DateTime::get_Now());
    Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(get_MyDir() + u"Document.docx", get_ArtifactsDir() + u"File.DetectDigitalSignatures.docx", certificateHolder, signOptions);
    
    // Use a new FileFormatInstance to confirm that it is signed.
    info = Aspose::Words::FileFormatUtil::DetectFileFormat(get_ArtifactsDir() + u"File.DetectDigitalSignatures.docx");
    
    ASSERT_TRUE(info->get_HasDigitalSignature());
    
    // We can load and access the signatures of a signed document in a collection like this.
    ASSERT_EQ(1, Aspose::Words::DigitalSignatures::DigitalSignatureUtil::LoadSignatures(get_ArtifactsDir() + u"File.DetectDigitalSignatures.docx")->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFile, DetectDigitalSignatures)
{
    s_instance->DetectDigitalSignatures();
}

} // namespace gtest_test

void ExFile::SaveToDetectedFileFormat()
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
        System::SharedPtr<System::IO::FileStream> docStream = System::IO::File::OpenRead(get_MyDir() + u"Word document with missing file extension");
        System::SharedPtr<Aspose::Words::FileFormatInfo> info = Aspose::Words::FileFormatUtil::DetectFileFormat(docStream);
        Aspose::Words::LoadFormat loadFormat = info->get_LoadFormat();
        
        ASSERT_EQ(Aspose::Words::LoadFormat::Doc, loadFormat);
        
        // Below are two methods of converting a LoadFormat to its corresponding SaveFormat.
        // 1 -  Get the file extension string for the LoadFormat, then get the corresponding SaveFormat from that string:
        System::String fileExtension = Aspose::Words::FileFormatUtil::LoadFormatToExtension(loadFormat);
        Aspose::Words::SaveFormat saveFormat = Aspose::Words::FileFormatUtil::ExtensionToSaveFormat(fileExtension);
        
        // 2 -  Convert the LoadFormat directly to its SaveFormat:
        saveFormat = Aspose::Words::FileFormatUtil::LoadFormatToSaveFormat(loadFormat);
        
        // Load a document from the stream, and then save it to the automatically detected file extension.
        auto doc = System::MakeObject<Aspose::Words::Document>(docStream);
        
        ASSERT_EQ(u".doc", Aspose::Words::FileFormatUtil::SaveFormatToExtension(saveFormat));
        
        doc->Save(get_ArtifactsDir() + u"File.SaveToDetectedFileFormat" + Aspose::Words::FileFormatUtil::SaveFormatToExtension(saveFormat));
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFile, SaveToDetectedFileFormat)
{
    s_instance->SaveToDetectedFileFormat();
}

} // namespace gtest_test

void ExFile::DetectFileFormat_SaveFormatToLoadFormat()
{
    //ExStart
    //ExFor:FileFormatUtil.SaveFormatToLoadFormat(SaveFormat)
    //ExSummary:Shows how to convert a save format to its corresponding load format.
    ASSERT_EQ(Aspose::Words::LoadFormat::Html, Aspose::Words::FileFormatUtil::SaveFormatToLoadFormat(Aspose::Words::SaveFormat::Html));
    
    // Some file types can have documents saved to, but not loaded from using Aspose.Words.
    // If we attempt to convert a save format of such a type to a load format, an exception will be thrown.
    ASSERT_THROW(static_cast<std::function<void()>>([]() -> void
    {
        Aspose::Words::FileFormatUtil::SaveFormatToLoadFormat(Aspose::Words::SaveFormat::Jpeg);
    })(), System::ArgumentException);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExFile, DetectFileFormat_SaveFormatToLoadFormat)
{
    s_instance->DetectFileFormat_SaveFormatToLoadFormat();
}

} // namespace gtest_test

void ExFile::ExtractImages()
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
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Images.docx");
    
    // Get the collection of shapes from the document,
    // and save the image data of every shape with an image as a file to the local file system.
    System::SharedPtr<Aspose::Words::NodeCollection> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);
    
    ASSERT_EQ(9, shapes->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> s)>>([](System::SharedPtr<Aspose::Words::Node> s) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Drawing::Shape>(s))->get_HasImage();
    }))));
    
    int32_t imageIndex = 0;
    for (auto&& shape : System::IterateOver(shapes->LINQ_OfType<System::SharedPtr<Aspose::Words::Drawing::Shape> >()))
    {
        if (shape->get_HasImage())
        {
            // The image data of shapes may contain images of many possible image formats. 
            // We can determine a file extension for each image automatically, based on its format.
            System::String imageFileName = System::String::Format(u"File.ExtractImages.{0}{1}", imageIndex, Aspose::Words::FileFormatUtil::ImageTypeToExtension(shape->get_ImageData()->get_ImageType()));
            shape->get_ImageData()->Save(get_ArtifactsDir() + imageFileName);
            imageIndex++;
        }
    }
    //ExEnd
    
    ASSERT_EQ(9, System::IO::Directory::GetFiles(get_ArtifactsDir())->LINQ_Count(static_cast<System::Func<System::String, bool>>(static_cast<std::function<bool(System::String s)>>([](System::String s) -> bool
    {
        return System::Text::RegularExpressions::Regex::IsMatch(s, u"^.+\\.(jpeg|png|emf|wmf)$") && s.StartsWith(get_ArtifactsDir() + u"File.ExtractImages");
    }))));
}

namespace gtest_test
{

TEST_F(ExFile, ExtractImages)
{
    s_instance->ExtractImages();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
