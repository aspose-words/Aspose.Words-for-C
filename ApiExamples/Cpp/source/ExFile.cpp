// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/io/file_stream.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/date_time.h>
#include <system/console.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
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
#include <Aspose.Words.Cpp/Model/Document/SignOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/LoadFormat.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>
#include <Aspose.Words.Cpp/Model/Document/FileCorruptedException.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/Document/CertificateHolder.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
namespace ApiExamples {

namespace gtest_test
{

class ExFile : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExFile> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExFile>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExFile> ExFile::s_instance;

} // namespace gtest_test

void ExFile::CatchFileCorruptedException()
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
        System::Console::WriteLine(e->get_Message());
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
    // 'DetectFileFormat' not working on a non-html files
    SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(MyDir + u"Document.docx");
    ASSERT_EQ(Aspose::Words::LoadFormat::Docx, info->get_LoadFormat());
    ASSERT_TRUE(info->get_Encoding() == nullptr);

    // This time the property will not be null
    info = FileFormatUtil::DetectFileFormat(MyDir + u"Document.html");
    ASSERT_EQ(Aspose::Words::LoadFormat::Html, info->get_LoadFormat());
    ASSERT_FALSE(info->get_Encoding() == nullptr);
    //ExEnd
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
    //ExSummary:Shows how to find the corresponding Aspose load/save format from an IANA content type string.
    // Trying to search for a SaveFormat with a simple string will not work
    try
    {
        ASSERT_EQ(Aspose::Words::SaveFormat::Jpeg, FileFormatUtil::ContentTypeToSaveFormat(u"jpeg"));
    }
    catch (System::ArgumentException& e)
    {
        System::Console::WriteLine(e->get_Message());
    }

    // The convertion methods only accept official IANA type names, which are all listed here:
    //      https://www.iana.org/assignments/media-types/media-types.xhtml
    // Note that if a corresponding SaveFormat or LoadFormat for a type from that list does not exist in the Aspose enums,
    // converting will raise an exception just like in the code above

    // File types that can be saved to but not opened as documents will not have corresponding load formats
    // Attempting to convert them to load formats will raise an exception
    ASSERT_EQ(Aspose::Words::SaveFormat::Jpeg, FileFormatUtil::ContentTypeToSaveFormat(u"image/jpeg"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Png, FileFormatUtil::ContentTypeToSaveFormat(u"image/png"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Tiff, FileFormatUtil::ContentTypeToSaveFormat(u"image/tiff"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Gif, FileFormatUtil::ContentTypeToSaveFormat(u"image/gif"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Emf, FileFormatUtil::ContentTypeToSaveFormat(u"image/x-emf"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Xps, FileFormatUtil::ContentTypeToSaveFormat(u"application/vnd.ms-xpsdocument"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Pdf, FileFormatUtil::ContentTypeToSaveFormat(u"application/pdf"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Svg, FileFormatUtil::ContentTypeToSaveFormat(u"image/svg+xml"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Epub, FileFormatUtil::ContentTypeToSaveFormat(u"application/epub+zip"));

    // File types that can both be loaded and saved have corresponding load and save formats
    ASSERT_EQ(Aspose::Words::LoadFormat::Doc, FileFormatUtil::ContentTypeToLoadFormat(u"application/msword"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Doc, FileFormatUtil::ContentTypeToSaveFormat(u"application/msword"));

    ASSERT_EQ(Aspose::Words::LoadFormat::Docx, FileFormatUtil::ContentTypeToLoadFormat(u"application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Docx, FileFormatUtil::ContentTypeToSaveFormat(u"application/vnd.openxmlformats-officedocument.wordprocessingml.document"));

    ASSERT_EQ(Aspose::Words::LoadFormat::Text, FileFormatUtil::ContentTypeToLoadFormat(u"text/plain"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Text, FileFormatUtil::ContentTypeToSaveFormat(u"text/plain"));

    ASSERT_EQ(Aspose::Words::LoadFormat::Rtf, FileFormatUtil::ContentTypeToLoadFormat(u"application/rtf"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Rtf, FileFormatUtil::ContentTypeToSaveFormat(u"application/rtf"));

    ASSERT_EQ(Aspose::Words::LoadFormat::Html, FileFormatUtil::ContentTypeToLoadFormat(u"text/html"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Html, FileFormatUtil::ContentTypeToSaveFormat(u"text/html"));

    ASSERT_EQ(Aspose::Words::LoadFormat::Mhtml, FileFormatUtil::ContentTypeToLoadFormat(u"multipart/related"));
    ASSERT_EQ(Aspose::Words::SaveFormat::Mhtml, FileFormatUtil::ContentTypeToSaveFormat(u"multipart/related"));
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
    auto doc = MakeObject<Document>();

    // Save it as an encrypted .odt
    auto saveOptions = MakeObject<OdtSaveOptions>(Aspose::Words::SaveFormat::Odt);
    saveOptions->set_Password(u"MyPassword");

    doc->Save(ArtifactsDir + u"File.DetectDocumentEncryption.odt", saveOptions);

    // Create a FileFormatInfo object for this document
    SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(ArtifactsDir + u"File.DetectDocumentEncryption.odt");

    // Verify the file type of our document and its encryption status
    ASSERT_EQ(u".odt", FileFormatUtil::LoadFormatToExtension(info->get_LoadFormat()));
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
    // Use a FileFormatInfo instance to verify that a document is not digitally signed
    SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(MyDir + u"Document.docx");

    ASSERT_EQ(u".docx", FileFormatUtil::LoadFormatToExtension(info->get_LoadFormat()));
    ASSERT_FALSE(info->get_HasDigitalSignature());

    // Sign the document
    SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw", nullptr);
    DigitalSignatureUtil::Sign(MyDir + u"Document.docx", ArtifactsDir + u"File.DetectDigitalSignatures.docx", certificateHolder, [&]{ auto tmp_0 = MakeObject<SignOptions>(); tmp_0->set_SignTime(System::DateTime::get_Now()); return tmp_0; }());

    // Use a new FileFormatInstance to confirm that it is signed
    info = FileFormatUtil::DetectFileFormat(ArtifactsDir + u"File.DetectDigitalSignatures.docx");

    ASSERT_TRUE(info->get_HasDigitalSignature());

    // The signatures can then be accessed like this
    ASSERT_EQ(1, DigitalSignatureUtil::LoadSignatures(ArtifactsDir + u"File.DetectDigitalSignatures.docx")->get_Count());
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
    //ExSummary:Shows how to use the FileFormatUtil methods to detect the format of a document without any extension and save it with the correct file extension.
    // Load the document without a file extension into a stream and use the DetectFileFormat method to detect it's format
    // These are both times where you might need extract the file format as it's not visible
    // The file format of this document is actually ".doc"
    SharedPtr<System::IO::FileStream> docStream = System::IO::File::OpenRead(MyDir + u"Word document with missing file extension");
    SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(docStream);

    // Retrieve the LoadFormat of the document
    LoadFormat loadFormat = info->get_LoadFormat();

    // Let's show the different methods of converting LoadFormat enumerations to SaveFormat enumerations
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
    //ExSummary:Shows how to use the FileFormatUtil class and to convert a SaveFormat enumeration into the corresponding LoadFormat enumeration.
    // Define the SaveFormat enumeration to convert
    const SaveFormat saveFormat = Aspose::Words::SaveFormat::Html;
    // Convert the SaveFormat enumeration to LoadFormat enumeration
    LoadFormat loadFormat = FileFormatUtil::SaveFormatToLoadFormat(saveFormat);
    System::Console::WriteLine(String(u"The converted LoadFormat is: ") + FileFormatUtil::LoadFormatToExtension(loadFormat));
    //ExEnd

    ASSERT_EQ(u".html", FileFormatUtil::SaveFormatToExtension(saveFormat));
    ASSERT_EQ(u".html", FileFormatUtil::LoadFormatToExtension(loadFormat));
}

namespace gtest_test
{

TEST_F(ExFile, DetectFileFormat_SaveFormatToLoadFormat)
{
    s_instance->DetectFileFormat_SaveFormatToLoadFormat();
}

} // namespace gtest_test

void ExFile::ExtractImagesToFiles()
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

    SharedPtr<NodeCollection> shapes = doc->GetChildNodes(Aspose::Words::NodeType::Shape, true);

    ASSERT_EQ(9, shapes->LINQ_Count([](SharedPtr<Node> s) { return (System::DynamicCast<Shape>(s))->get_HasImage(); }));

    int imageIndex = 0;
    for (auto shape : System::IterateOver(shapes->LINQ_OfType<SharedPtr<Shape> >()))
    {
        if (shape->get_HasImage())
        {
            String imageFileName = String::Format(u"File.ExtractImagesToFiles.{0}{1}",imageIndex,FileFormatUtil::ImageTypeToExtension(shape->get_ImageData()->get_ImageType()));
            shape->get_ImageData()->Save(ArtifactsDir + imageFileName);
            imageIndex++;
        }
    }
    //ExEnd

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(String s)> _local_func_1 = [](String s)
    {
        return System::Text::RegularExpressions::Regex::IsMatch(s, u"^.+\\.(jpeg|png|emf|wmf)$") && s.StartsWith(ArtifactsDir + u"File.ExtractImagesToFiles");
    };

    ASSERT_EQ(9, System::IO::Directory::GetFiles(ArtifactsDir)->LINQ_Count(static_cast<System::Func<String, bool>>(_local_func_1)));
}

namespace gtest_test
{

TEST_F(ExFile, ExtractImagesToFiles)
{
    s_instance->ExtractImagesToFiles();
}

} // namespace gtest_test

} // namespace ApiExamples
