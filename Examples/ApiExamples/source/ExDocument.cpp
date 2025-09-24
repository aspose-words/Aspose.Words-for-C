// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocument.h"

#include <testing/test_predicates.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match_collection.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/memory_stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_not_found_exception.h>
#include <system/io/file_mode.h>
#include <system/io/file_access.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/globalization/culture_info.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/enum_helpers.h>
#include <system/details/dispose_guard.h>
#include <system/date_time.h>
#include <system/convert.h>
#include <system/collections/list.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <security/cryptography/x509_certificates/x509_certificate_2.h>
#include <security/cryptography/x509_certificates/x500_distinguished_name.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/size.h>
#include <drawing/image.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaProject.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaModuleCollection.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaModule.h>
#include <Aspose.Words.Cpp/Rendering/ThumbnailGeneratingOptions.h>
#include <Aspose.Words.Cpp/Rendering/PageInfo.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtensionReference.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtensionProperty.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtensionBinding.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtension.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/TaskPane.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Enums/WebExtensionStoreType.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Enums/WebExtensionBindingType.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Enums/TaskPaneDockState.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Collections/WebExtensionReferenceCollection.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Collections/WebExtensionPropertyCollection.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Collections/WebExtensionBindingCollection.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Collections/TaskPaneCollection.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Settings/WriteProtection.h>
#include <Aspose.Words.Cpp/Model/Settings/JustificationMode.h>
#include <Aspose.Words.Cpp/Model/Settings/HyphenationOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Orientation.h>
#include <Aspose.Words.Cpp/Model/Sections/Margins.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/XamlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionCollection.h>
#include <Aspose.Words.Cpp/Model/Revisions/Revision.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomPart.h>
#include <Aspose.Words.Cpp/Model/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLabel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Framesets/FramesetCollection.h>
#include <Aspose.Words.Cpp/Model/Framesets/Frameset.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteOptions.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Fonts/MemoryFontSource.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontInfoCollection.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Drawing/Watermark/WatermarkType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Watermark/WatermarkLayout.h>
#include <Aspose.Words.Cpp/Model/Drawing/Watermark/Watermark.h>
#include <Aspose.Words.Cpp/Model/Drawing/Watermark/TextWatermarkOptions.h>
#include <Aspose.Words.Cpp/Model/Drawing/Watermark/ImageWatermarkOptions.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/SubDocument.h>
#include <Aspose.Words.Cpp/Model/Document/ProtectionType.h>
#include <Aspose.Words.Cpp/Model/Document/PageExtractOptions.h>
#include <Aspose.Words.Cpp/Model/Document/LoadFormat.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Model/Document/ImportFormatOptions.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureType.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignature.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/CertificateHolder.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutOptions.h>
#include <Aspose.Words.Cpp/Layout/Public/CommentDisplayMode.h>

#include "TestUtil.h"
#include "DocumentHelper.h"


using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Rendering;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
using namespace Aspose::Words::Vba;
using namespace Aspose::Words::WebExtensions;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1054267473u, ::Aspose::Words::ApiExamples::ExDocument::HandleNodeChangingFontChanger, ThisTypeBaseTypesInfo);

void ExDocument::HandleNodeChangingFontChanger::NodeInserted(System::SharedPtr<Aspose::Words::NodeChangingArgs> args)
{
    mLog->AppendLine(System::String::Format(u"\tType:\t{0}", args->get_Node()->get_NodeType()));
    mLog->AppendLine(System::String::Format(u"\tHash:\t{0}", System::ObjectExt::GetHashCode(args->get_Node())));
    
    if (args->get_Node()->get_NodeType() == Aspose::Words::NodeType::Run)
    {
        System::SharedPtr<Aspose::Words::Font> font = (System::ExplicitCast<Aspose::Words::Run>(args->get_Node()))->get_Font();
        mLog->Append(System::String::Format(u"\tFont:\tChanged from \"{0}\" {1}pt", font->get_Name(), font->get_Size()));
        
        font->set_Size(24);
        font->set_Name(u"Arial");
        
        mLog->AppendLine(System::String::Format(u" to \"{0}\" {1}pt", font->get_Name(), font->get_Size()));
        mLog->AppendLine(System::String::Format(u"\tContents:\n\t\t\"{0}\"", args->get_Node()->GetText()));
    }
}

void ExDocument::HandleNodeChangingFontChanger::NodeInserting(System::SharedPtr<Aspose::Words::NodeChangingArgs> args)
{
    mLog->AppendLine(System::String::Format(u"\n{0:dd/MM/yyyy HH:mm:ss:fff}\tNode insertion:", System::DateTime::get_Now()));
}

void ExDocument::HandleNodeChangingFontChanger::NodeRemoved(System::SharedPtr<Aspose::Words::NodeChangingArgs> args)
{
    mLog->AppendLine(System::String::Format(u"\tType:\t{0}", args->get_Node()->get_NodeType()));
    mLog->AppendLine(System::String::Format(u"\tHash code:\t{0}", System::ObjectExt::GetHashCode(args->get_Node())));
}

void ExDocument::HandleNodeChangingFontChanger::NodeRemoving(System::SharedPtr<Aspose::Words::NodeChangingArgs> args)
{
    mLog->AppendLine(System::String::Format(u"\n{0:dd/MM/yyyy HH:mm:ss:fff}\tNode removal:", System::DateTime::get_Now()));
}

System::String ExDocument::HandleNodeChangingFontChanger::GetLog()
{
    return mLog->ToString();
}

ExDocument::HandleNodeChangingFontChanger::HandleNodeChangingFontChanger()
    : mLog(System::MakeObject<System::Text::StringBuilder>())
{
}


RTTI_INFO_IMPL_HASH(4178396843u, ::Aspose::Words::ApiExamples::ExDocument, ThisTypeBaseTypesInfo);

void ExDocument::TestFontChangeViaCallback(System::String log)
{
    ASSERT_EQ(10, System::Text::RegularExpressions::Regex::Matches(log, u"insertion")->get_Count());
    ASSERT_EQ(5, System::Text::RegularExpressions::Regex::Matches(log, u"removal")->get_Count());
}

void ExDocument::TestDocPackageCustomParts(System::SharedPtr<Aspose::Words::Markup::CustomPartCollection> parts)
{
    ASSERT_EQ(3, parts->get_Count());
    
    ASSERT_EQ(u"/payload/payload_on_package.test", parts->idx_get(0)->get_Name());
    ASSERT_EQ(u"mytest/somedata", parts->idx_get(0)->get_ContentType());
    ASSERT_EQ(u"http://mytest.payload.internal", parts->idx_get(0)->get_RelationshipType());
    ASPOSE_ASSERT_EQ(false, parts->idx_get(0)->get_IsExternal());
    ASSERT_EQ(18, parts->idx_get(0)->get_Data()->get_Length());
    
    ASSERT_EQ(u"http://www.aspose.com/Images/aspose-logo.jpg", parts->idx_get(1)->get_Name());
    ASSERT_EQ(u"", parts->idx_get(1)->get_ContentType());
    ASSERT_EQ(u"http://mytest.payload.external", parts->idx_get(1)->get_RelationshipType());
    ASPOSE_ASSERT_EQ(true, parts->idx_get(1)->get_IsExternal());
    ASSERT_EQ(0, parts->idx_get(1)->get_Data()->get_Length());
    
    ASSERT_EQ(u"http://www.aspose.com/Images/aspose-logo.jpg", parts->idx_get(2)->get_Name());
    ASSERT_EQ(u"", parts->idx_get(2)->get_ContentType());
    ASSERT_EQ(u"http://mytest.payload.external", parts->idx_get(2)->get_RelationshipType());
    ASPOSE_ASSERT_EQ(true, parts->idx_get(2)->get_IsExternal());
    ASSERT_EQ(0, parts->idx_get(2)->get_Data()->get_Length());
}


namespace gtest_test
{

class ExDocument : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExDocument> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExDocument>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExDocument> ExDocument::s_instance;

} // namespace gtest_test

void ExDocument::CreateSimpleDocument()
{
    //ExStart:CreateSimpleDocument
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:DocumentBase.Document
    //ExFor:Document.#ctor()
    //ExSummary:Shows how to create simple document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // New Document objects by default come with the minimal set of nodes
    // required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
    doc->AppendChild<System::SharedPtr<Aspose::Words::Section>>(System::MakeObject<Aspose::Words::Section>(doc))->AppendChild<System::SharedPtr<Aspose::Words::Body>>(System::MakeObject<Aspose::Words::Body>(doc))->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(System::MakeObject<Aspose::Words::Paragraph>(doc))->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!"));
    //ExEnd:CreateSimpleDocument
}

namespace gtest_test
{

TEST_F(ExDocument, CreateSimpleDocument)
{
    s_instance->CreateSimpleDocument();
}

} // namespace gtest_test

void ExDocument::Constructor()
{
    //ExStart
    //ExFor:Document.#ctor()
    //ExFor:Document.#ctor(String,LoadOptions)
    //ExSummary:Shows how to create and load documents.
    // There are two ways of creating a Document object using Aspose.Words.
    // 1 -  Create a blank document:
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // New Document objects by default come with the minimal set of nodes
    // required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!"));
    
    // 2 -  Load a document that exists in the local file system:
    doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    // Loaded documents will have contents that we can access and edit.
    ASSERT_EQ(u"Hello World!", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText().Trim());
    
    // Some operations that need to occur during loading, such as using a password to decrypt a document,
    // can be done by passing a LoadOptions object when loading the document.
    doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Encrypted.docx", System::MakeObject<Aspose::Words::Loading::LoadOptions>(u"docPassword"));
    
    ASSERT_EQ(u"Test encrypted document.", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, Constructor)
{
    s_instance->Constructor();
}

} // namespace gtest_test

void ExDocument::LoadFromStream()
{
    //ExStart
    //ExFor:Document.#ctor(Stream)
    //ExSummary:Shows how to load a document using a stream.
    {
        System::SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(get_MyDir() + u"Document.docx");
        auto doc = System::MakeObject<Aspose::Words::Document>(stream);
        
        ASSERT_EQ(u"Hello World!\r\rHello Word!\r\r\rHello World!", doc->GetText().Trim());
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, LoadFromStream)
{
    s_instance->LoadFromStream();
}

} // namespace gtest_test

void ExDocument::ConvertToPdf()
{
    //ExStart
    //ExFor:Document.#ctor(String)
    //ExFor:Document.Save(String)
    //ExSummary:Shows how to open a document and convert it to .PDF.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    doc->Save(get_ArtifactsDir() + u"Document.ConvertToPdf.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, ConvertToPdf)
{
    s_instance->ConvertToPdf();
}

} // namespace gtest_test

void ExDocument::SaveToImageStream()
{
    //ExStart
    //ExFor:Document.Save(Stream, SaveFormat)
    //ExSummary:Shows how to save a document to an image via stream, and then read the image from that stream.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Times New Roman");
    builder->get_Font()->set_Size(24);
    builder->Writeln(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, SaveToImageStream)
{
    s_instance->SaveToImageStream();
}

} // namespace gtest_test

void ExDocument::DetectMobiDocumentFormat()
{
    System::SharedPtr<Aspose::Words::FileFormatInfo> info = Aspose::Words::FileFormatUtil::DetectFileFormat(get_MyDir() + u"Document.mobi");
    ASSERT_EQ(info->get_LoadFormat(), Aspose::Words::LoadFormat::Mobi);
}

namespace gtest_test
{

TEST_F(ExDocument, DetectMobiDocumentFormat)
{
    s_instance->DetectMobiDocumentFormat();
}

} // namespace gtest_test

void ExDocument::OpenFromStreamWithBaseUri()
{
    //ExStart
    //ExFor:Document.#ctor(Stream,LoadOptions)
    //ExFor:LoadOptions.#ctor
    //ExFor:LoadOptions.BaseUri
    //ExFor:ShapeBase.IsImage
    //ExSummary:Shows how to open an HTML document with images from a stream using a base URI.
    {
        System::SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(get_MyDir() + u"Document.html");
        // Pass the URI of the base folder while loading it
        // so that any images with relative URIs in the HTML document can be found.
        auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
        loadOptions->set_BaseUri(get_ImageDir());
        
        auto doc = System::MakeObject<Aspose::Words::Document>(stream, loadOptions);
        
        // Verify that the first shape of the document contains a valid image.
        auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
        
        ASSERT_TRUE(shape->get_IsImage());
        ASSERT_FALSE(System::TestTools::IsNull(shape->get_ImageData()->get_ImageBytes()));
        ASSERT_NEAR(32.0, Aspose::Words::ConvertUtil::PointToPixel(shape->get_Width()), 0.01);
        ASSERT_NEAR(32.0, Aspose::Words::ConvertUtil::PointToPixel(shape->get_Height()), 0.01);
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, OpenFromStreamWithBaseUri)
{
    s_instance->OpenFromStreamWithBaseUri();
}

} // namespace gtest_test

void ExDocument::LoadEncrypted()
{
    //ExStart
    //ExFor:Document.#ctor(Stream,LoadOptions)
    //ExFor:Document.#ctor(String,LoadOptions)
    //ExFor:LoadOptions
    //ExFor:LoadOptions.#ctor(String)
    //ExSummary:Shows how to load an encrypted Microsoft Word document.
    System::SharedPtr<Aspose::Words::Document> doc;
    
    // Aspose.Words throw an exception if we try to open an encrypted document without its password.
    ASSERT_THROW(static_cast<std::function<void()>>([&doc]() -> void
    {
        doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Encrypted.docx");
    })(), Aspose::Words::IncorrectPasswordException);
    
    // When loading such a document, the password is passed to the document's constructor using a LoadOptions object.
    auto options = System::MakeObject<Aspose::Words::Loading::LoadOptions>(u"docPassword");
    
    // There are two ways of loading an encrypted document with a LoadOptions object.
    // 1 -  Load the document from the local file system by filename:
    doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Encrypted.docx", options);
    ASSERT_EQ(u"Test encrypted document.", doc->GetText().Trim());
    //ExSkip
    
    // 2 -  Load the document from a stream:
    {
        System::SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(get_MyDir() + u"Encrypted.docx");
        doc = System::MakeObject<Aspose::Words::Document>(stream, options);
        ASSERT_EQ(u"Test encrypted document.", doc->GetText().Trim());
        //ExSkip
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, LoadEncrypted)
{
    s_instance->LoadEncrypted();
}

} // namespace gtest_test

void ExDocument::NotSupportedWarning()
{
    //ExStart
    //ExFor:WarningInfoCollection.Count
    //ExFor:WarningInfoCollection.Item(Int32)
    //ExSummary:Shows how to get warnings about unsupported formats.
    auto warnings = System::MakeObject<Aspose::Words::WarningInfoCollection>();
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_WarningCallback(warnings);
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"FB2 document.fb2", loadOptions);
    
    ASSERT_EQ(u"The original file load format is FB2, which is not supported by Aspose.Words. The file is loaded as an XML document.", warnings->idx_get(0)->get_Description());
    ASSERT_EQ(1, warnings->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, NotSupportedWarning)
{
    s_instance->NotSupportedWarning();
}

} // namespace gtest_test

void ExDocument::TempFolder()
{
    //ExStart
    //ExFor:LoadOptions.TempFolder
    //ExSummary:Shows how to load a document using temporary files.
    // Note that such an approach can reduce memory usage but degrades speed
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_TempFolder(u"C:\\TempFolder\\");
    
    // Ensure that the directory exists and load
    System::IO::Directory::CreateDirectory_(loadOptions->get_TempFolder());
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx", loadOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, TempFolder)
{
    s_instance->TempFolder();
}

} // namespace gtest_test

void ExDocument::ConvertToHtml()
{
    //ExStart
    //ExFor:Document.Save(String,SaveFormat)
    //ExFor:SaveFormat
    //ExSummary:Shows how to convert from DOCX to HTML format.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    doc->Save(get_ArtifactsDir() + u"Document.ConvertToHtml.html", Aspose::Words::SaveFormat::Html);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, ConvertToHtml)
{
    s_instance->ConvertToHtml();
}

} // namespace gtest_test

void ExDocument::ConvertToMhtml()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    doc->Save(get_ArtifactsDir() + u"Document.ConvertToMhtml.mht");
}

namespace gtest_test
{

TEST_F(ExDocument, ConvertToMhtml)
{
    s_instance->ConvertToMhtml();
}

} // namespace gtest_test

void ExDocument::ConvertToTxt()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    doc->Save(get_ArtifactsDir() + u"Document.ConvertToTxt.txt");
}

namespace gtest_test
{

TEST_F(ExDocument, ConvertToTxt)
{
    s_instance->ConvertToTxt();
}

} // namespace gtest_test

void ExDocument::ConvertToEpub()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    doc->Save(get_ArtifactsDir() + u"Document.ConvertToEpub.epub");
}

namespace gtest_test
{

TEST_F(ExDocument, ConvertToEpub)
{
    s_instance->ConvertToEpub();
}

} // namespace gtest_test

void ExDocument::SaveToStream()
{
    //ExStart
    //ExFor:Document.Save(Stream,SaveFormat)
    //ExSummary:Shows how to save a document to a stream.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    {
        auto dstStream = System::MakeObject<System::IO::MemoryStream>();
        doc->Save(dstStream, Aspose::Words::SaveFormat::Docx);
        
        // Verify that the stream contains the document.
        ASSERT_EQ(u"Hello World!\r\rHello Word!\r\r\rHello World!", System::MakeObject<Aspose::Words::Document>(dstStream)->GetText().Trim());
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, SaveToStream)
{
    s_instance->SaveToStream();
}

} // namespace gtest_test

void ExDocument::FontChangeViaCallback()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Set the node changing callback to custom implementation,
    // then add/remove nodes to get it to generate a log.
    auto callback = System::MakeObject<Aspose::Words::ApiExamples::ExDocument::HandleNodeChangingFontChanger>();
    doc->set_NodeChangingCallback(callback);
    
    builder->Writeln(u"Hello world!");
    builder->Writeln(u"Hello again!");
    builder->InsertField(u" HYPERLINK \"https://www.google.com/\" ");
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 300, 300);
    
    doc->get_Range()->get_Fields()->idx_get(0)->Remove();
    
    std::cout << callback->GetLog() << std::endl;
    TestFontChangeViaCallback(callback->GetLog());
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocument, FontChangeViaCallback)
{
    s_instance->FontChangeViaCallback();
}

} // namespace gtest_test

void ExDocument::AppendDocument()
{
    //ExStart
    //ExFor:Document.AppendDocument(Document, ImportFormatMode)
    //ExSummary:Shows how to append a document to the end of another document.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>();
    srcDoc->get_FirstSection()->get_Body()->AppendParagraph(u"Source document text. ");
    
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    dstDoc->get_FirstSection()->get_Body()->AppendParagraph(u"Destination document text. ");
    
    // Append the source document to the destination document while preserving its formatting,
    // then save the source document to the local file system.
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting);
    ASSERT_EQ(2, dstDoc->get_Sections()->get_Count());
    //ExSkip
    
    dstDoc->Save(get_ArtifactsDir() + u"Document.AppendDocument.docx");
    //ExEnd
    
    System::String outDocText = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.AppendDocument.docx")->GetText();
    
    ASSERT_TRUE(outDocText.StartsWith(dstDoc->GetText()));
    ASSERT_TRUE(outDocText.EndsWith(srcDoc->GetText()));
}

namespace gtest_test
{

TEST_F(ExDocument, AppendDocument)
{
    s_instance->AppendDocument();
}

} // namespace gtest_test

void ExDocument::AppendDocumentFromAutomation()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // We should call this method to clear this document of any existing content.
    doc->RemoveAllChildren();
    
    const int32_t recordCount = 5;
    for (int32_t i = 1; i <= recordCount; i++)
    {
        auto srcDoc = System::MakeObject<Aspose::Words::Document>();
        
        ASSERT_THROW(static_cast<std::function<void()>>([]() -> void
        {
            System::MakeObject<Aspose::Words::Document>(System::String(u"C:\\DetailsList.doc"));
        })(), System::IO::FileNotFoundException);
        
        // Append the source document at the end of the destination document.
        doc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::UseDestinationStyles);
        
        // Automation required you to insert a new section break at this point, however, in Aspose.Words we
        // do not need to do anything here as the appended document is imported as separate sections already
        
        // Unlink all headers/footers in this section from the previous section headers/footers
        // if this is the second document or above being appended.
        if (i > 1)
        {
            ASSERT_THROW(static_cast<std::function<void()>>([&doc, &i]() -> void
            {
                doc->get_Sections()->idx_get(i)->get_HeadersFooters()->LinkToPrevious(false);
            })(), System::NullReferenceException);
        }
    }
}

namespace gtest_test
{

TEST_F(ExDocument, AppendDocumentFromAutomation)
{
    s_instance->AppendDocumentFromAutomation();
}

} // namespace gtest_test

void ExDocument::ImportList(bool isKeepSourceNumbering)
{
    //ExStart
    //ExFor:ImportFormatOptions.KeepSourceNumbering
    //ExSummary:Shows how to import a document with numbered lists.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"List source.docx");
    auto dstDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"List destination.docx");
    
    ASSERT_EQ(4, dstDoc->get_Lists()->get_Count());
    
    auto options = System::MakeObject<Aspose::Words::ImportFormatOptions>();
    
    // If there is a clash of list styles, apply the list format of the source document.
    // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
    // Set the "KeepSourceNumbering" property to "true" import all clashing
    // list style numbering with the same appearance that it had in the source document.
    options->set_KeepSourceNumbering(isKeepSourceNumbering);
    
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting, options);
    dstDoc->UpdateListLabels();
    
    ASSERT_EQ(isKeepSourceNumbering ? 5 : 4, dstDoc->get_Lists()->get_Count());
    //ExEnd
}

namespace gtest_test
{

using ExDocument_ImportList_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocument::ImportList)>::type;

struct ExDocument_ImportList : public ExDocument, public Aspose::Words::ApiExamples::ExDocument, public ::testing::WithParamInterface<ExDocument_ImportList_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExDocument_ImportList, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ImportList(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_ImportList, ::testing::ValuesIn(ExDocument_ImportList::TestCases()));

} // namespace gtest_test

void ExDocument::KeepSourceNumberingSameListIds()
{
    //ExStart
    //ExFor:ImportFormatOptions.KeepSourceNumbering
    //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how resolve a clash when importing documents that have lists with the same list definition identifier.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"List with the same definition identifier - source.docx");
    auto dstDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"List with the same definition identifier - destination.docx");
    
    // Set the "KeepSourceNumbering" property to "true" to apply a different list definition ID
    // to identical styles as Aspose.Words imports them into destination documents.
    auto importFormatOptions = System::MakeObject<Aspose::Words::ImportFormatOptions>();
    importFormatOptions->set_KeepSourceNumbering(true);
    
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::UseDestinationStyles, importFormatOptions);
    dstDoc->UpdateListLabels();
    //ExEnd
    
    System::String paraText = dstDoc->get_Sections()->idx_get(1)->get_Body()->get_LastParagraph()->GetText();
    
    ASSERT_TRUE(paraText.StartsWith(u"13->13")) << (paraText);
    ASSERT_EQ(u"1.", dstDoc->get_Sections()->idx_get(1)->get_Body()->get_LastParagraph()->get_ListLabel()->get_LabelString());
}

namespace gtest_test
{

TEST_F(ExDocument, KeepSourceNumberingSameListIds)
{
    s_instance->KeepSourceNumberingSameListIds();
}

} // namespace gtest_test

void ExDocument::MergePastedLists()
{
    //ExStart
    //ExFor:ImportFormatOptions.MergePastedLists
    //ExSummary:Shows how to merge lists from a documents.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"List item.docx");
    auto dstDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"List destination.docx");
    
    auto options = System::MakeObject<Aspose::Words::ImportFormatOptions>();
    options->set_MergePastedLists(true);
    
    // Set the "MergePastedLists" property to "true" pasted lists will be merged with surrounding lists.
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::UseDestinationStyles, options);
    
    dstDoc->Save(get_ArtifactsDir() + u"Document.MergePastedLists.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, MergePastedLists)
{
    s_instance->MergePastedLists();
}

} // namespace gtest_test

void ExDocument::ForceCopyStyles()
{
    //ExStart
    //ExFor:ImportFormatOptions.ForceCopyStyles
    //ExSummary:Shows how to copy source styles with unique names forcibly.
    // Both documents contain MyStyle1 and MyStyle2, MyStyle3 exists only in a source document.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Styles source.docx");
    auto dstDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Styles destination.docx");
    
    auto options = System::MakeObject<Aspose::Words::ImportFormatOptions>();
    options->set_ForceCopyStyles(true);
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting, options);
    
    System::SharedPtr<Aspose::Words::ParagraphCollection> paras = dstDoc->get_Sections()->idx_get(1)->get_Body()->get_Paragraphs();
    
    ASSERT_EQ(paras->idx_get(0)->get_ParagraphFormat()->get_Style()->get_Name(), u"MyStyle1_0");
    ASSERT_EQ(paras->idx_get(1)->get_ParagraphFormat()->get_Style()->get_Name(), u"MyStyle2_0");
    ASSERT_EQ(paras->idx_get(2)->get_ParagraphFormat()->get_Style()->get_Name(), u"MyStyle3");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, ForceCopyStyles)
{
    s_instance->ForceCopyStyles();
}

} // namespace gtest_test

void ExDocument::AdjustSentenceAndWordSpacing()
{
    //ExStart
    //ExFor:ImportFormatOptions.AdjustSentenceAndWordSpacing
    //ExSummary:Shows how to adjust sentence and word spacing automatically.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>();
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(srcDoc);
    builder->Write(u"Dolor sit amet.");
    
    builder = System::MakeObject<Aspose::Words::DocumentBuilder>(dstDoc);
    builder->Write(u"Lorem ipsum.");
    
    auto options = System::MakeObject<Aspose::Words::ImportFormatOptions>();
    options->set_AdjustSentenceAndWordSpacing(true);
    builder->InsertDocument(srcDoc, Aspose::Words::ImportFormatMode::UseDestinationStyles, options);
    
    ASSERT_EQ(u"Lorem ipsum. Dolor sit amet.", dstDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, AdjustSentenceAndWordSpacing)
{
    s_instance->AdjustSentenceAndWordSpacing();
}

} // namespace gtest_test

void ExDocument::ValidateIndividualDocumentSignatures()
{
    //ExStart
    //ExFor:CertificateHolder.Certificate
    //ExFor:Document.DigitalSignatures
    //ExFor:DigitalSignature
    //ExFor:DigitalSignatureCollection
    //ExFor:DigitalSignature.IsValid
    //ExFor:DigitalSignature.Comments
    //ExFor:DigitalSignature.SignTime
    //ExFor:DigitalSignature.SignatureType
    //ExSummary:Shows how to validate and display information about each signature in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Digitally signed.docx");
    
    for (auto&& signature : doc->get_DigitalSignatures())
    {
        std::cout << System::String::Format(u"{0} signature: ", (signature->get_IsValid() ? System::String(u"Valid") : System::String(u"Invalid"))) << std::endl;
        std::cout << System::String::Format(u"\tReason:\t{0}", signature->get_Comments()) << std::endl;
        std::cout << System::String::Format(u"\tType:\t{0}", signature->get_SignatureType()) << std::endl;
        std::cout << System::String::Format(u"\tSign time:\t{0}", signature->get_SignTime()) << std::endl;
        std::cout << System::String::Format(u"\tSubject name:\t{0}", signature->get_CertificateHolder()->get_Certificate()->get_SubjectName()) << std::endl;
        std::cout << System::String::Format(u"\tIssuer name:\t{0}", signature->get_CertificateHolder()->get_Certificate()->get_IssuerName()->get_Name()) << std::endl;
        std::cout << std::endl;
    }
    //ExEnd
    
    ASSERT_EQ(1, doc->get_DigitalSignatures()->get_Count());
    
    System::SharedPtr<Aspose::Words::DigitalSignatures::DigitalSignature> digitalSig = doc->get_DigitalSignatures()->idx_get(0);
    
    ASSERT_TRUE(digitalSig->get_IsValid());
    ASSERT_EQ(u"Test Sign", digitalSig->get_Comments());
    ASSERT_EQ(u"XmlDsig", System::ObjectExt::ToString(digitalSig->get_SignatureType()));
    ASSERT_TRUE(digitalSig->get_CertificateHolder()->get_Certificate()->get_Subject().Contains(u"Aspose Pty Ltd"));
    ASSERT_TRUE(digitalSig->get_CertificateHolder()->get_Certificate()->get_IssuerName()->get_Name() != nullptr && digitalSig->get_CertificateHolder()->get_Certificate()->get_IssuerName()->get_Name().Contains(u"VeriSign"));
}

namespace gtest_test
{

TEST_F(ExDocument, ValidateIndividualDocumentSignatures)
{
    s_instance->ValidateIndividualDocumentSignatures();
}

} // namespace gtest_test

void ExDocument::DigitalSignature()
{
    //ExStart
    //ExFor:DigitalSignature.CertificateHolder
    //ExFor:DigitalSignature.IssuerName
    //ExFor:DigitalSignature.SubjectName
    //ExFor:DigitalSignatureCollection
    //ExFor:DigitalSignatureCollection.IsValid
    //ExFor:DigitalSignatureCollection.Count
    //ExFor:DigitalSignatureCollection.Item(Int32)
    //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder)
    //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder)
    //ExFor:DigitalSignatureType
    //ExFor:Document.DigitalSignatures
    //ExSummary:Shows how to sign documents with X.509 certificates.
    // Verify that a document is not signed.
    ASSERT_FALSE(Aspose::Words::FileFormatUtil::DetectFileFormat(get_MyDir() + u"Document.docx")->get_HasDigitalSignature());
    
    // Create a CertificateHolder object from a PKCS12 file, which we will use to sign the document.
    System::SharedPtr<Aspose::Words::DigitalSignatures::CertificateHolder> certificateHolder = Aspose::Words::DigitalSignatures::CertificateHolder::Create(get_MyDir() + u"morzal.pfx", u"aw", nullptr);
    
    // There are two ways of saving a signed copy of a document to the local file system:
    // 1 - Designate a document by a local system filename and save a signed copy at a location specified by another filename.
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_SignTime(System::DateTime::get_Now());
    Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(get_MyDir() + u"Document.docx", get_ArtifactsDir() + u"Document.DigitalSignature.docx", certificateHolder, signOptions);
    
    ASSERT_TRUE(Aspose::Words::FileFormatUtil::DetectFileFormat(get_ArtifactsDir() + u"Document.DigitalSignature.docx")->get_HasDigitalSignature());
    
    // 2 - Take a document from a stream and save a signed copy to another stream.
    {
        auto inDoc = System::MakeObject<System::IO::FileStream>(get_MyDir() + u"Document.docx", System::IO::FileMode::Open);
        {
            auto outDoc = System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + u"Document.DigitalSignature.docx", System::IO::FileMode::Create);
            Aspose::Words::DigitalSignatures::DigitalSignatureUtil::Sign(inDoc, outDoc, certificateHolder);
        }
    }
    
    ASSERT_TRUE(Aspose::Words::FileFormatUtil::DetectFileFormat(get_ArtifactsDir() + u"Document.DigitalSignature.docx")->get_HasDigitalSignature());
    
    // Please verify that all of the document's digital signatures are valid and check their details.
    auto signedDoc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.DigitalSignature.docx");
    System::SharedPtr<Aspose::Words::DigitalSignatures::DigitalSignatureCollection> digitalSignatureCollection = signedDoc->get_DigitalSignatures();
    
    ASSERT_TRUE(digitalSignatureCollection->get_IsValid());
    ASSERT_EQ(1, digitalSignatureCollection->get_Count());
    ASSERT_EQ(Aspose::Words::DigitalSignatures::DigitalSignatureType::XmlDsig, digitalSignatureCollection->idx_get(0)->get_SignatureType());
    ASSERT_EQ(u"CN=Morzal.Me", signedDoc->get_DigitalSignatures()->idx_get(0)->get_IssuerName());
    ASSERT_EQ(u"CN=Morzal.Me", signedDoc->get_DigitalSignatures()->idx_get(0)->get_SubjectName());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, DigitalSignature)
{
    s_instance->DigitalSignature();
}

} // namespace gtest_test

void ExDocument::SignatureValue()
{
    //ExStart
    //ExFor:DigitalSignature.SignatureValue
    //ExSummary:Shows how to get a digital signature value from a digitally signed document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Digitally signed.docx");
    
    for (auto&& digitalSignature : doc->get_DigitalSignatures())
    {
        System::String signatureValue = System::Convert::ToBase64String(digitalSignature->get_SignatureValue());
        ASSERT_EQ(System::String(u"K1cVLLg2kbJRAzT5WK+m++G8eEO+l7S+5ENdjMxxTXkFzGUfvwxREuJdSFj9AbD") + u"MhnGvDURv9KEhC25DDF1al8NRVR71TF3CjHVZXpYu7edQS5/yLw/k5CiFZzCp1+MmhOdYPcVO+Fm" + u"+9fKr2iNLeyYB+fgEeZHfTqTFM2WwAqo=", signatureValue);
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, SignatureValue)
{
    s_instance->SignatureValue();
}

} // namespace gtest_test

void ExDocument::AppendAllDocumentsInFolder()
{
    //ExStart
    //ExFor:Document.AppendDocument(Document, ImportFormatMode)
    //ExSummary:Shows how to append all the documents in a folder to the end of a template document.
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(dstDoc);
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading1);
    builder->Writeln(u"Template Document");
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Normal);
    builder->Writeln(u"Some content here");
    ASSERT_EQ(5, dstDoc->get_Styles()->get_Count());
    //ExSkip
    ASSERT_EQ(1, dstDoc->get_Sections()->get_Count());
    //ExSkip
    
    // Append all unencrypted documents with the .doc extension
    // from our local file system directory to the base document.
    System::SharedPtr<System::Collections::Generic::List<System::String>> docFiles = System::IO::Directory::GetFiles(get_MyDir(), u"*.doc")->LINQ_Where(static_cast<System::Func<System::String, bool>>(static_cast<std::function<bool(System::String item)>>([](System::String item) -> bool
    {
        return item.EndsWith(u".doc");
    })))->LINQ_ToList();
    for (auto&& fileName : docFiles)
    {
        System::SharedPtr<Aspose::Words::FileFormatInfo> info = Aspose::Words::FileFormatUtil::DetectFileFormat(fileName);
        if (info->get_IsEncrypted())
        {
            continue;
        }
        
        auto srcDoc = System::MakeObject<Aspose::Words::Document>(fileName);
        dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::UseDestinationStyles);
    }
    
    dstDoc->Save(get_ArtifactsDir() + u"Document.AppendAllDocumentsInFolder.doc");
    //ExEnd
    
    ASSERT_EQ(7, dstDoc->get_Styles()->get_Count());
    ASSERT_EQ(10, dstDoc->get_Sections()->get_Count());
}

namespace gtest_test
{

TEST_F(ExDocument, AppendAllDocumentsInFolder)
{
    s_instance->AppendAllDocumentsInFolder();
}

} // namespace gtest_test

void ExDocument::JoinRunsWithSameFormatting()
{
    //ExStart
    //ExFor:Document.JoinRunsWithSameFormatting
    //ExSummary:Shows how to join runs in a document to reduce unneeded runs.
    // Open a document that contains adjacent runs of text with identical formatting,
    // which commonly occurs if we edit the same paragraph multiple times in Microsoft Word.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // If any number of these runs are adjacent with identical formatting,
    // then the document may be simplified.
    ASSERT_EQ(317, doc->GetChildNodes(Aspose::Words::NodeType::Run, true)->get_Count());
    
    // Combine such runs with this method and verify the number of run joins that will take place.
    ASSERT_EQ(121, doc->JoinRunsWithSameFormatting());
    
    // The number of joins and the number of runs we have after the join
    // should add up the number of runs we had initially.
    ASSERT_EQ(196, doc->GetChildNodes(Aspose::Words::NodeType::Run, true)->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, JoinRunsWithSameFormatting)
{
    s_instance->JoinRunsWithSameFormatting();
}

} // namespace gtest_test

void ExDocument::DefaultTabStop()
{
    //ExStart
    //ExFor:Document.DefaultTabStop
    //ExFor:ControlChar.Tab
    //ExFor:ControlChar.TabChar
    //ExSummary:Shows how to set a custom interval for tab stop positions.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Set tab stops to appear every 72 points (1 inch).
    builder->get_Document()->set_DefaultTabStop(72);
    
    // Each tab character snaps the text after it to the next closest tab stop position.
    builder->Writeln(System::String(u"Hello") + Aspose::Words::ControlChar::Tab() + u"World!");
    builder->Writeln(System::String(u"Hello") + Aspose::Words::ControlChar::TabChar + u"World!");
    //ExEnd
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    ASPOSE_ASSERT_EQ(72, doc->get_DefaultTabStop());
}

namespace gtest_test
{

TEST_F(ExDocument, DefaultTabStop)
{
    s_instance->DefaultTabStop();
}

} // namespace gtest_test

void ExDocument::CloneDocument()
{
    //ExStart
    //ExFor:Document.Clone
    //ExSummary:Shows how to deep clone a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Hello world!");
    
    // Cloning will produce a new document with the same contents as the original,
    // but with a unique copy of each of the original document's nodes.
    System::SharedPtr<Aspose::Words::Document> clone = doc->Clone();
    
    ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->GetText(), clone->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Text());
    ASSERT_NE(System::ObjectExt::GetHashCode(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)), System::ObjectExt::GetHashCode(clone->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, CloneDocument)
{
    s_instance->CloneDocument();
}

} // namespace gtest_test

void ExDocument::DocumentGetTextToString()
{
    //ExStart
    //ExFor:CompositeNode.GetText
    //ExFor:Node.ToString(SaveFormat)
    //ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->InsertField(u"MERGEFIELD Field");
    
    // GetText will retrieve the visible text as well as field codes and special characters.
    ASSERT_EQ(u"\u0013MERGEFIELD Field\u0014«Field»\u0015", doc->GetText().Trim());
    
    // ToString will give us the document's appearance if saved to a passed save format.
    ASSERT_EQ(u"«Field»", doc->ToString(Aspose::Words::SaveFormat::Text).Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, DocumentGetTextToString)
{
    s_instance->DocumentGetTextToString();
}

} // namespace gtest_test

void ExDocument::ProtectUnprotect()
{
    //ExStart
    //ExFor:Document.Protect(ProtectionType,String)
    //ExFor:Document.ProtectionType
    //ExFor:Document.Unprotect
    //ExFor:Document.Unprotect(String)
    //ExSummary:Shows how to protect and unprotect a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->Protect(Aspose::Words::ProtectionType::ReadOnly, u"password");
    
    ASSERT_EQ(Aspose::Words::ProtectionType::ReadOnly, doc->get_ProtectionType());
    
    // If we open this document with Microsoft Word intending to edit it,
    // we will need to apply the password to get through the protection.
    doc->Save(get_ArtifactsDir() + u"Document.Protect.docx");
    
    // Note that the protection only applies to Microsoft Word users opening our document.
    // We have not encrypted the document in any way, and we do not need the password to open and edit it programmatically.
    auto protectedDoc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.Protect.docx");
    
    ASSERT_EQ(Aspose::Words::ProtectionType::ReadOnly, protectedDoc->get_ProtectionType());
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(protectedDoc);
    builder->Writeln(u"Text added to a protected document.");
    ASSERT_EQ(u"Text added to a protected document.", protectedDoc->get_Range()->get_Text().Trim());
    //ExSkip
    
    // There are two ways of removing protection from a document.
    // 1 - With no password:
    doc->Unprotect();
    
    ASSERT_EQ(Aspose::Words::ProtectionType::NoProtection, doc->get_ProtectionType());
    
    doc->Protect(Aspose::Words::ProtectionType::ReadOnly, u"NewPassword");
    
    ASSERT_EQ(Aspose::Words::ProtectionType::ReadOnly, doc->get_ProtectionType());
    
    doc->Unprotect(u"WrongPassword");
    
    ASSERT_EQ(Aspose::Words::ProtectionType::ReadOnly, doc->get_ProtectionType());
    
    // 2 - With the correct password:
    doc->Unprotect(u"NewPassword");
    
    ASSERT_EQ(Aspose::Words::ProtectionType::NoProtection, doc->get_ProtectionType());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, ProtectUnprotect)
{
    s_instance->ProtectUnprotect();
}

} // namespace gtest_test

void ExDocument::DocumentEnsureMinimum()
{
    //ExStart
    //ExFor:Document.EnsureMinimum
    //ExSummary:Shows how to ensure that a document contains the minimal set of nodes required for editing its contents.
    // A newly created document contains one child Section, which includes one child Body and one child Paragraph.
    // We can edit the document body's contents by adding nodes such as Runs or inline Shapes to that paragraph.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::NodeCollection> nodes = doc->GetChildNodes(Aspose::Words::NodeType::Any, true);
    
    ASSERT_EQ(Aspose::Words::NodeType::Section, nodes->idx_get(0)->get_NodeType());
    ASPOSE_ASSERT_EQ(doc, nodes->idx_get(0)->get_ParentNode());
    
    ASSERT_EQ(Aspose::Words::NodeType::Body, nodes->idx_get(1)->get_NodeType());
    ASPOSE_ASSERT_EQ(nodes->idx_get(0), nodes->idx_get(1)->get_ParentNode());
    
    ASSERT_EQ(Aspose::Words::NodeType::Paragraph, nodes->idx_get(2)->get_NodeType());
    ASPOSE_ASSERT_EQ(nodes->idx_get(1), nodes->idx_get(2)->get_ParentNode());
    
    // This is the minimal set of nodes that we need to be able to edit the document.
    // We will no longer be able to edit the document if we remove any of them.
    doc->RemoveAllChildren();
    
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
    
    // Call this method to make sure that the document has at least those three nodes so we can edit it again.
    doc->EnsureMinimum();
    
    ASSERT_EQ(Aspose::Words::NodeType::Section, nodes->idx_get(0)->get_NodeType());
    ASSERT_EQ(Aspose::Words::NodeType::Body, nodes->idx_get(1)->get_NodeType());
    ASSERT_EQ(Aspose::Words::NodeType::Paragraph, nodes->idx_get(2)->get_NodeType());
    
    (System::ExplicitCast<Aspose::Words::Paragraph>(nodes->idx_get(2)))->get_Runs()->Add(System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!"));
    //ExEnd
    
    ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocument, DocumentEnsureMinimum)
{
    s_instance->DocumentEnsureMinimum();
}

} // namespace gtest_test

void ExDocument::RemoveMacrosFromDocument()
{
    //ExStart
    //ExFor:Document.RemoveMacros
    //ExSummary:Shows how to remove all macros from a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Macro.docm");
    
    ASSERT_TRUE(doc->get_HasMacros());
    ASSERT_EQ(u"Project", doc->get_VbaProject()->get_Name());
    
    // Remove the document's VBA project, along with all its macros.
    doc->RemoveMacros();
    
    ASSERT_FALSE(doc->get_HasMacros());
    ASSERT_TRUE(System::TestTools::IsNull(doc->get_VbaProject()));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, RemoveMacrosFromDocument)
{
    s_instance->RemoveMacrosFromDocument();
}

} // namespace gtest_test

void ExDocument::GetPageCount()
{
    //ExStart
    //ExFor:Document.PageCount
    //ExSummary:Shows how to count the number of pages in the document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Page 1");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Write(u"Page 2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Write(u"Page 3");
    
    // Verify the expected page count of the document.
    ASSERT_EQ(3, doc->get_PageCount());
    
    // Getting the PageCount property invoked the document's page layout to calculate the value.
    // This operation will not need to be re-done when rendering the document to a fixed page save format,
    // such as .pdf. So you can save some time, especially with more complex documents.
    doc->Save(get_ArtifactsDir() + u"Document.GetPageCount.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, GetPageCount)
{
    s_instance->GetPageCount();
}

} // namespace gtest_test

void ExDocument::GetUpdatedPageProperties()
{
    //ExStart
    //ExFor:Document.UpdateWordCount()
    //ExFor:Document.UpdateWordCount(Boolean)
    //ExFor:BuiltInDocumentProperties.Characters
    //ExFor:BuiltInDocumentProperties.Words
    //ExFor:BuiltInDocumentProperties.Paragraphs
    //ExFor:BuiltInDocumentProperties.Lines
    //ExSummary:Shows how to update all list labels in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(System::String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") + u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    builder->Write(System::String(u"Ut enim ad minim veniam, ") + u"quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");
    
    // Aspose.Words does not track document metrics like these in real time.
    ASSERT_EQ(0, doc->get_BuiltInDocumentProperties()->get_Characters());
    ASSERT_EQ(0, doc->get_BuiltInDocumentProperties()->get_Words());
    ASSERT_EQ(1, doc->get_BuiltInDocumentProperties()->get_Paragraphs());
    ASSERT_EQ(1, doc->get_BuiltInDocumentProperties()->get_Lines());
    
    // To get accurate values for three of these properties, we will need to update them manually.
    doc->UpdateWordCount();
    
    ASSERT_EQ(196, doc->get_BuiltInDocumentProperties()->get_Characters());
    ASSERT_EQ(36, doc->get_BuiltInDocumentProperties()->get_Words());
    ASSERT_EQ(2, doc->get_BuiltInDocumentProperties()->get_Paragraphs());
    
    // For the line count, we will need to call a specific overload of the updating method.
    ASSERT_EQ(1, doc->get_BuiltInDocumentProperties()->get_Lines());
    
    doc->UpdateWordCount(true);
    
    ASSERT_EQ(4, doc->get_BuiltInDocumentProperties()->get_Lines());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, GetUpdatedPageProperties)
{
    s_instance->GetUpdatedPageProperties();
}

} // namespace gtest_test

void ExDocument::GetOriginalFileInfo()
{
    //ExStart
    //ExFor:Document.OriginalFileName
    //ExFor:Document.OriginalLoadFormat
    //ExSummary:Shows how to retrieve details of a document's load operation.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    ASSERT_EQ(get_MyDir() + u"Document.docx", doc->get_OriginalFileName());
    ASSERT_EQ(Aspose::Words::LoadFormat::Docx, doc->get_OriginalLoadFormat());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, GetOriginalFileInfo)
{
    s_instance->GetOriginalFileInfo();
}

} // namespace gtest_test

void ExDocument::FootnoteColumns()
{
    //ExStart
    //ExFor:FootnoteOptions
    //ExFor:FootnoteOptions.Columns
    //ExSummary:Shows how to split the footnote section into a given number of columns.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Footnotes and endnotes.docx");
    ASSERT_EQ(0, doc->get_FootnoteOptions()->get_Columns());
    //ExSkip
    
    doc->get_FootnoteOptions()->set_Columns(2);
    doc->Save(get_ArtifactsDir() + u"Document.FootnoteColumns.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.FootnoteColumns.docx");
    
    ASSERT_EQ(2, doc->get_FirstSection()->get_PageSetup()->get_FootnoteOptions()->get_Columns());
}

namespace gtest_test
{

TEST_F(ExDocument, FootnoteColumns)
{
    s_instance->FootnoteColumns();
}

} // namespace gtest_test

void ExDocument::RemoveExternalSchemaReferences()
{
    //ExStart
    //ExFor:Document.RemoveExternalSchemaReferences
    //ExSummary:Shows how to remove all external XML schema references from a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"External XML schema.docx");
    
    doc->RemoveExternalSchemaReferences();
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, RemoveExternalSchemaReferences)
{
    s_instance->RemoveExternalSchemaReferences();
}

} // namespace gtest_test

void ExDocument::UpdateThumbnail()
{
    //ExStart
    //ExFor:Document.UpdateThumbnail()
    //ExFor:Document.UpdateThumbnail(ThumbnailGeneratingOptions)
    //ExFor:ThumbnailGeneratingOptions
    //ExFor:ThumbnailGeneratingOptions.GenerateFromFirstPage
    //ExFor:ThumbnailGeneratingOptions.ThumbnailSize
    //ExSummary:Shows how to update a document's thumbnail.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Hello world!");
    builder->InsertImage(get_ImageDir() + u"Logo.jpg");
    
    // There are two ways of setting a thumbnail image when saving a document to .epub.
    // 1 -  Use the document's first page:
    doc->UpdateThumbnail();
    doc->Save(get_ArtifactsDir() + u"Document.UpdateThumbnail.FirstPage.epub");
    
    // 2 -  Use the first image found in the document:
    auto options = System::MakeObject<Aspose::Words::Rendering::ThumbnailGeneratingOptions>();
    ASPOSE_ASSERT_EQ(System::Drawing::Size(600, 900), options->get_ThumbnailSize());
    //ExSkip
    ASSERT_TRUE(options->get_GenerateFromFirstPage());
    //ExSkip
    options->set_ThumbnailSize(System::Drawing::Size(400, 400));
    options->set_GenerateFromFirstPage(false);
    
    doc->UpdateThumbnail(options);
    doc->Save(get_ArtifactsDir() + u"Document.UpdateThumbnail.FirstImage.epub");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, UpdateThumbnail)
{
    s_instance->UpdateThumbnail();
}

} // namespace gtest_test

void ExDocument::HyphenationOptions()
{
    //ExStart
    //ExFor:Document.HyphenationOptions
    //ExFor:HyphenationOptions
    //ExFor:HyphenationOptions.AutoHyphenation
    //ExFor:HyphenationOptions.ConsecutiveHyphenLimit
    //ExFor:HyphenationOptions.HyphenationZone
    //ExFor:HyphenationOptions.HyphenateCaps
    //ExSummary:Shows how to configure automatic hyphenation.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Size(24);
    builder->Writeln(System::String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") + u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    
    doc->get_HyphenationOptions()->set_AutoHyphenation(true);
    doc->get_HyphenationOptions()->set_ConsecutiveHyphenLimit(2);
    doc->get_HyphenationOptions()->set_HyphenationZone(720);
    doc->get_HyphenationOptions()->set_HyphenateCaps(true);
    
    doc->Save(get_ArtifactsDir() + u"Document.HyphenationOptions.docx");
    //ExEnd
    
    ASPOSE_ASSERT_EQ(true, doc->get_HyphenationOptions()->get_AutoHyphenation());
    ASSERT_EQ(2, doc->get_HyphenationOptions()->get_ConsecutiveHyphenLimit());
    ASSERT_EQ(720, doc->get_HyphenationOptions()->get_HyphenationZone());
    ASPOSE_ASSERT_EQ(true, doc->get_HyphenationOptions()->get_HyphenateCaps());
    
    ASSERT_TRUE(Aspose::Words::ApiExamples::DocumentHelper::CompareDocs(get_ArtifactsDir() + u"Document.HyphenationOptions.docx", get_GoldsDir() + u"Document.HyphenationOptions Gold.docx"));
}

namespace gtest_test
{

TEST_F(ExDocument, HyphenationOptions)
{
    s_instance->HyphenationOptions();
}

} // namespace gtest_test

void ExDocument::HyphenationOptionsDefaultValues()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    
    ASPOSE_ASSERT_EQ(false, doc->get_HyphenationOptions()->get_AutoHyphenation());
    ASSERT_EQ(0, doc->get_HyphenationOptions()->get_ConsecutiveHyphenLimit());
    ASSERT_EQ(360, doc->get_HyphenationOptions()->get_HyphenationZone());
    // 0.25 inch
    ASPOSE_ASSERT_EQ(true, doc->get_HyphenationOptions()->get_HyphenateCaps());
}

namespace gtest_test
{

TEST_F(ExDocument, HyphenationOptionsDefaultValues)
{
    s_instance->HyphenationOptionsDefaultValues();
}

} // namespace gtest_test

void ExDocument::HyphenationZoneException()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    ASSERT_THROW(static_cast<std::function<void()>>([&doc]() -> void
    {
        doc->get_HyphenationOptions()->set_HyphenationZone(0);
    })(), System::ArgumentOutOfRangeException);
}

namespace gtest_test
{

TEST_F(ExDocument, HyphenationZoneException)
{
    s_instance->HyphenationZoneException();
}

} // namespace gtest_test

void ExDocument::OoxmlComplianceVersion()
{
    //ExStart
    //ExFor:Document.Compliance
    //ExSummary:Shows how to read a loaded document's Open Office XML compliance version.
    // The compliance version varies between documents created by different versions of Microsoft Word.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.doc");
    ASSERT_EQ(doc->get_Compliance(), Aspose::Words::Saving::OoxmlCompliance::Ecma376_2006);
    
    doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    ASSERT_EQ(doc->get_Compliance(), Aspose::Words::Saving::OoxmlCompliance::Iso29500_2008_Transitional);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, OoxmlComplianceVersion)
{
    s_instance->OoxmlComplianceVersion();
}

} // namespace gtest_test

void ExDocument::ImageSaveOptions()
{
    //ExStart
    //ExFor:Document.Save(String, SaveOptions)
    //ExFor:SaveOptions.UseAntiAliasing
    //ExFor:SaveOptions.UseHighQualityRendering
    //ExSummary:Shows how to improve the quality of a rendered document with SaveOptions.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Size(60);
    builder->Writeln(u"Some text.");
    
    System::SharedPtr<Aspose::Words::Saving::SaveOptions> options = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Jpeg);
    ASSERT_FALSE(options->get_UseAntiAliasing());
    //ExSkip
    ASSERT_FALSE(options->get_UseHighQualityRendering());
    //ExSkip
    
    doc->Save(get_ArtifactsDir() + u"Document.ImageSaveOptions.Default.jpg", options);
    
    options->set_UseAntiAliasing(true);
    options->set_UseHighQualityRendering(true);
    
    doc->Save(get_ArtifactsDir() + u"Document.ImageSaveOptions.HighQuality.jpg", options);
    //ExEnd
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(794, 1122, get_ArtifactsDir() + u"Document.ImageSaveOptions.Default.jpg");
    Aspose::Words::ApiExamples::TestUtil::VerifyImage(794, 1122, get_ArtifactsDir() + u"Document.ImageSaveOptions.HighQuality.jpg");
}

namespace gtest_test
{

TEST_F(ExDocument, ImageSaveOptions)
{
    s_instance->ImageSaveOptions();
}

} // namespace gtest_test

void ExDocument::Cleanup()
{
    //ExStart
    //ExFor:Document.Cleanup
    //ExSummary:Shows how to remove unused custom styles from a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    doc->get_Styles()->Add(Aspose::Words::StyleType::List, u"MyListStyle1");
    doc->get_Styles()->Add(Aspose::Words::StyleType::List, u"MyListStyle2");
    doc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"MyParagraphStyle1");
    doc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"MyParagraphStyle2");
    
    // Combined with the built-in styles, the document now has eight styles.
    // A custom style counts as "used" while applied to some part of the document,
    // which means that the four styles we added are currently unused.
    ASSERT_EQ(8, doc->get_Styles()->get_Count());
    
    // Apply a custom character style, and then a custom list style. Doing so will mark the styles as "used".
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->get_Font()->set_Style(doc->get_Styles()->idx_get(u"MyParagraphStyle1"));
    builder->Writeln(u"Hello world!");
    
    System::SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(doc->get_Styles()->idx_get(u"MyListStyle1"));
    builder->get_ListFormat()->set_List(list);
    builder->Writeln(u"Item 1");
    builder->Writeln(u"Item 2");
    
    doc->Cleanup();
    
    ASSERT_EQ(6, doc->get_Styles()->get_Count());
    
    // Removing every node that a custom style is applied to marks it as "unused" again.
    // Run the Cleanup method again to remove them.
    doc->get_FirstSection()->get_Body()->RemoveAllChildren();
    doc->Cleanup();
    
    ASSERT_EQ(4, doc->get_Styles()->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, Cleanup)
{
    s_instance->Cleanup();
}

} // namespace gtest_test

void ExDocument::AutomaticallyUpdateStyles()
{
    //ExStart
    //ExFor:Document.AutomaticallyUpdateStyles
    //ExSummary:Shows how to attach a template to a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Microsoft Word documents by default come with an attached template called "Normal.dotm".
    // There is no default template for blank Aspose.Words documents.
    ASSERT_EQ(System::String::Empty, doc->get_AttachedTemplate());
    
    // Attach a template, then set the flag to apply style changes
    // within the template to styles in our document.
    doc->set_AttachedTemplate(get_MyDir() + u"Business brochure.dotx");
    doc->set_AutomaticallyUpdateStyles(true);
    
    doc->Save(get_ArtifactsDir() + u"Document.AutomaticallyUpdateStyles.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.AutomaticallyUpdateStyles.docx");
    
    ASSERT_TRUE(doc->get_AutomaticallyUpdateStyles());
    ASSERT_EQ(get_MyDir() + u"Business brochure.dotx", doc->get_AttachedTemplate());
    ASSERT_TRUE(System::IO::File::Exists(doc->get_AttachedTemplate()));
}

namespace gtest_test
{

TEST_F(ExDocument, AutomaticallyUpdateStyles)
{
    s_instance->AutomaticallyUpdateStyles();
}

} // namespace gtest_test

void ExDocument::DefaultTemplate()
{
    //ExStart
    //ExFor:Document.AttachedTemplate
    //ExFor:Document.AutomaticallyUpdateStyles
    //ExFor:SaveOptions.CreateSaveOptions(String)
    //ExFor:SaveOptions.DefaultTemplate
    //ExSummary:Shows how to set a default template for documents that do not have attached templates.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Enable automatic style updating, but do not attach a template document.
    doc->set_AutomaticallyUpdateStyles(true);
    
    ASSERT_EQ(System::String::Empty, doc->get_AttachedTemplate());
    
    // Since there is no template document, the document had nowhere to track style changes.
    // Use a SaveOptions object to automatically set a template
    // if a document that we are saving does not have one.
    System::SharedPtr<Aspose::Words::Saving::SaveOptions> options = Aspose::Words::Saving::SaveOptions::CreateSaveOptions(u"Document.DefaultTemplate.docx");
    options->set_DefaultTemplate(get_MyDir() + u"Business brochure.dotx");
    
    doc->Save(get_ArtifactsDir() + u"Document.DefaultTemplate.docx", options);
    //ExEnd
    
    ASSERT_TRUE(System::IO::File::Exists(options->get_DefaultTemplate()));
}

namespace gtest_test
{

TEST_F(ExDocument, DefaultTemplate)
{
    s_instance->DefaultTemplate();
}

} // namespace gtest_test

void ExDocument::UseSubstitutions()
{
    //ExStart
    //ExFor:FindReplaceOptions.#ctor()
    //ExFor:FindReplaceOptions.UseSubstitutions
    //ExFor:FindReplaceOptions.LegacyMode
    //ExSummary:Shows how to recognize and use substitutions within replacement patterns.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Jason gave money to Paul.");
    
    auto regex = System::MakeObject<System::Text::RegularExpressions::Regex>(u"([A-z]+) gave money to ([A-z]+)");
    
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    options->set_UseSubstitutions(true);
    
    // Using legacy mode does not support many advanced features, so we need to set it to 'false'.
    options->set_LegacyMode(false);
    
    doc->get_Range()->Replace(regex, u"$2 took money from $1", options);
    
    ASSERT_EQ(doc->GetText(), u"Paul took money from Jason.\f");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, UseSubstitutions)
{
    s_instance->UseSubstitutions();
}

} // namespace gtest_test

void ExDocument::SetInvalidateFieldTypes()
{
    //ExStart
    //ExFor:Document.NormalizeFieldTypes
    //ExFor:Range.NormalizeFieldTypes
    //ExSummary:Shows how to get the keep a field's type up to date with its field code.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Fields::Field> field = builder->InsertField(u"DATE", nullptr);
    
    // Aspose.Words automatically detects field types based on field codes.
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldDate, field->get_Type());
    
    // Manually change the raw text of the field, which determines the field code.
    auto fieldText = System::ExplicitCast<Aspose::Words::Run>(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetChildNodes(Aspose::Words::NodeType::Run, true)->idx_get(0));
    ASSERT_EQ(u"DATE", fieldText->get_Text());
    //ExSkip
    fieldText->set_Text(u"PAGE");
    
    // Changing the field code has changed this field to one of a different type,
    // but the field's type properties still display the old type.
    ASSERT_EQ(u"PAGE", field->GetFieldCode());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldDate, field->get_Type());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldDate, field->get_Start()->get_FieldType());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldDate, field->get_Separator()->get_FieldType());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldDate, field->get_End()->get_FieldType());
    
    // Update those properties with this method to display current value.
    doc->NormalizeFieldTypes();
    
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldPage, field->get_Type());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldPage, field->get_Start()->get_FieldType());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldPage, field->get_Separator()->get_FieldType());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldPage, field->get_End()->get_FieldType());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, SetInvalidateFieldTypes)
{
    s_instance->SetInvalidateFieldTypes();
}

} // namespace gtest_test

void ExDocument::LayoutOptionsHiddenText(bool showHiddenText)
{
    //ExStart
    //ExFor:Document.LayoutOptions
    //ExFor:LayoutOptions
    //ExFor:LayoutOptions.ShowHiddenText
    //ExSummary:Shows how to hide text in a rendered output document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    ASSERT_FALSE(doc->get_LayoutOptions()->get_ShowHiddenText());
    //ExSkip
    
    // Insert hidden text, then specify whether we wish to omit it from a rendered document.
    builder->Writeln(u"This text is not hidden.");
    builder->get_Font()->set_Hidden(true);
    builder->Writeln(u"This text is hidden.");
    
    doc->get_LayoutOptions()->set_ShowHiddenText(showHiddenText);
    
    doc->Save(get_ArtifactsDir() + u"Document.LayoutOptionsHiddenText.pdf");
    //ExEnd
}

namespace gtest_test
{

using ExDocument_LayoutOptionsHiddenText_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocument::LayoutOptionsHiddenText)>::type;

struct ExDocument_LayoutOptionsHiddenText : public ExDocument, public Aspose::Words::ApiExamples::ExDocument, public ::testing::WithParamInterface<ExDocument_LayoutOptionsHiddenText_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocument_LayoutOptionsHiddenText, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->LayoutOptionsHiddenText(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_LayoutOptionsHiddenText, ::testing::ValuesIn(ExDocument_LayoutOptionsHiddenText::TestCases()));

} // namespace gtest_test

void ExDocument::LayoutOptionsParagraphMarks(bool showParagraphMarks)
{
    //ExStart
    //ExFor:Document.LayoutOptions
    //ExFor:LayoutOptions
    //ExFor:LayoutOptions.ShowParagraphMarks
    //ExSummary:Shows how to show paragraph marks in a rendered output document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    ASSERT_FALSE(doc->get_LayoutOptions()->get_ShowParagraphMarks());
    //ExSkip
    
    // Add some paragraphs, then enable paragraph marks to show the ends of paragraphs
    // with a pilcrow (¶) symbol when we render the document.
    builder->Writeln(u"Hello world!");
    builder->Writeln(u"Hello again!");
    
    doc->get_LayoutOptions()->set_ShowParagraphMarks(showParagraphMarks);
    
    doc->Save(get_ArtifactsDir() + u"Document.LayoutOptionsParagraphMarks.pdf");
    //ExEnd
}

namespace gtest_test
{

using ExDocument_LayoutOptionsParagraphMarks_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocument::LayoutOptionsParagraphMarks)>::type;

struct ExDocument_LayoutOptionsParagraphMarks : public ExDocument, public Aspose::Words::ApiExamples::ExDocument, public ::testing::WithParamInterface<ExDocument_LayoutOptionsParagraphMarks_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocument_LayoutOptionsParagraphMarks, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->LayoutOptionsParagraphMarks(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_LayoutOptionsParagraphMarks, ::testing::ValuesIn(ExDocument_LayoutOptionsParagraphMarks::TestCases()));

} // namespace gtest_test

void ExDocument::UpdatePageLayout()
{
    //ExStart
    //ExFor:StyleCollection.Item(String)
    //ExFor:SectionCollection.Item(Int32)
    //ExFor:Document.UpdatePageLayout
    //ExFor:Margins
    //ExFor:PageSetup.Margins
    //ExSummary:Shows when to recalculate the page layout of the document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // Saving a document to PDF, to an image, or printing for the first time will automatically
    // cache the layout of the document within its pages.
    doc->Save(get_ArtifactsDir() + u"Document.UpdatePageLayout.1.pdf");
    
    // Modify the document in some way.
    doc->get_Styles()->idx_get(u"Normal")->get_Font()->set_Size(6);
    doc->get_Sections()->idx_get(0)->get_PageSetup()->set_Orientation(Aspose::Words::Orientation::Landscape);
    doc->get_Sections()->idx_get(0)->get_PageSetup()->set_Margins(Aspose::Words::Margins::Mirrored);
    
    // In the current version of Aspose.Words, modifying the document does not automatically rebuild
    // the cached page layout. If we wish for the cached layout
    // to stay up to date, we will need to update it manually.
    doc->UpdatePageLayout();
    
    doc->Save(get_ArtifactsDir() + u"Document.UpdatePageLayout.2.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, UpdatePageLayout)
{
    s_instance->UpdatePageLayout();
}

} // namespace gtest_test

void ExDocument::DocPackageCustomParts()
{
    //ExStart
    //ExFor:CustomPart
    //ExFor:CustomPart.ContentType
    //ExFor:CustomPart.RelationshipType
    //ExFor:CustomPart.IsExternal
    //ExFor:CustomPart.Data
    //ExFor:CustomPart.Name
    //ExFor:CustomPart.Clone
    //ExFor:CustomPartCollection
    //ExFor:CustomPartCollection.Add(CustomPart)
    //ExFor:CustomPartCollection.Clear
    //ExFor:CustomPartCollection.Clone
    //ExFor:CustomPartCollection.Count
    //ExFor:CustomPartCollection.GetEnumerator
    //ExFor:CustomPartCollection.Item(Int32)
    //ExFor:CustomPartCollection.RemoveAt(Int32)
    //ExFor:Document.PackageCustomParts
    //ExSummary:Shows how to access a document's arbitrary custom parts collection.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Custom parts OOXML package.docx");
    
    ASSERT_EQ(2, doc->get_PackageCustomParts()->get_Count());
    
    // Clone the second part, then add the clone to the collection.
    System::SharedPtr<Aspose::Words::Markup::CustomPart> clonedPart = doc->get_PackageCustomParts()->idx_get(1)->Clone();
    doc->get_PackageCustomParts()->Add(clonedPart);
    TestDocPackageCustomParts(doc->get_PackageCustomParts());
    //ExSkip
    
    ASSERT_EQ(3, doc->get_PackageCustomParts()->get_Count());
    
    // Enumerate over the collection and print every part.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Markup::CustomPart>>> enumerator = doc->get_PackageCustomParts()->GetEnumerator();
        int32_t index = 0;
        while (enumerator->MoveNext())
        {
            std::cout << System::String::Format(u"Part index {0}:", index) << std::endl;
            std::cout << System::String::Format(u"\tName:\t\t\t\t{0}", enumerator->get_Current()->get_Name()) << std::endl;
            std::cout << System::String::Format(u"\tContent type:\t\t{0}", enumerator->get_Current()->get_ContentType()) << std::endl;
            std::cout << System::String::Format(u"\tRelationship type:\t{0}", enumerator->get_Current()->get_RelationshipType()) << std::endl;
            std::cout << (enumerator->get_Current()->get_IsExternal() ? u"\tSourced from outside the document" : System::String::Format(u"\tStored within the document, length: {0} bytes", enumerator->get_Current()->get_Data()->get_Length())) << std::endl;
            index++;
        }
    }
    
    // We can remove elements from this collection individually, or all at once.
    doc->get_PackageCustomParts()->RemoveAt(2);
    
    ASSERT_EQ(2, doc->get_PackageCustomParts()->get_Count());
    
    doc->get_PackageCustomParts()->Clear();
    
    ASSERT_EQ(0, doc->get_PackageCustomParts()->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, DocPackageCustomParts)
{
    s_instance->DocPackageCustomParts();
}

} // namespace gtest_test

void ExDocument::ShadeFormData(bool useGreyShading)
{
    //ExStart
    //ExFor:Document.ShadeFormData
    //ExSummary:Shows how to apply gray shading to form fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    ASSERT_TRUE(doc->get_ShadeFormData());
    //ExSkip
    
    builder->Write(u"Hello world! ");
    builder->InsertTextInput(u"My form field", Aspose::Words::Fields::TextFormFieldType::Regular, u"", u"Text contents of form field, which are shaded in grey by default.", 0);
    
    // We can turn the grey shading off, so the bookmarked text will blend in with the other text.
    doc->set_ShadeFormData(useGreyShading);
    doc->Save(get_ArtifactsDir() + u"Document.ShadeFormData.docx");
    //ExEnd
}

namespace gtest_test
{

using ExDocument_ShadeFormData_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocument::ShadeFormData)>::type;

struct ExDocument_ShadeFormData : public ExDocument, public Aspose::Words::ApiExamples::ExDocument, public ::testing::WithParamInterface<ExDocument_ShadeFormData_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocument_ShadeFormData, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ShadeFormData(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_ShadeFormData, ::testing::ValuesIn(ExDocument_ShadeFormData::TestCases()));

} // namespace gtest_test

void ExDocument::VersionsCount()
{
    //ExStart
    //ExFor:Document.VersionsCount
    //ExSummary:Shows how to work with the versions count feature of older Microsoft Word documents.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Versions.doc");
    
    // We can read this property of a document, but we cannot preserve it while saving.
    ASSERT_EQ(4, doc->get_VersionsCount());
    
    doc->Save(get_ArtifactsDir() + u"Document.VersionsCount.doc");
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.VersionsCount.doc");
    
    ASSERT_EQ(0, doc->get_VersionsCount());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, VersionsCount)
{
    s_instance->VersionsCount();
}

} // namespace gtest_test

void ExDocument::WriteProtection()
{
    //ExStart
    //ExFor:Document.WriteProtection
    //ExFor:WriteProtection
    //ExFor:WriteProtection.IsWriteProtected
    //ExFor:WriteProtection.ReadOnlyRecommended
    //ExFor:WriteProtection.SetPassword(String)
    //ExFor:WriteProtection.ValidatePassword(String)
    //ExSummary:Shows how to protect a document with a password.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world! This document is protected.");
    ASSERT_FALSE(doc->get_WriteProtection()->get_IsWriteProtected());
    //ExSkip
    ASSERT_FALSE(doc->get_WriteProtection()->get_ReadOnlyRecommended());
    //ExSkip
    
    // Enter a password up to 15 characters in length, and then verify the document's protection status.
    doc->get_WriteProtection()->SetPassword(u"MyPassword");
    doc->get_WriteProtection()->set_ReadOnlyRecommended(true);
    
    ASSERT_TRUE(doc->get_WriteProtection()->get_IsWriteProtected());
    ASSERT_TRUE(doc->get_WriteProtection()->ValidatePassword(u"MyPassword"));
    
    // Protection does not prevent the document from being edited programmatically, nor does it encrypt the contents.
    doc->Save(get_ArtifactsDir() + u"Document.WriteProtection.docx");
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.WriteProtection.docx");
    
    ASSERT_TRUE(doc->get_WriteProtection()->get_IsWriteProtected());
    
    builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->MoveToDocumentEnd();
    builder->Writeln(u"Writing text in a protected document.");
    
    ASSERT_EQ(System::String(u"Hello world! This document is protected.") + u"\rWriting text in a protected document.", doc->GetText().Trim());
    //ExEnd
    ASSERT_TRUE(doc->get_WriteProtection()->get_ReadOnlyRecommended());
    ASSERT_TRUE(doc->get_WriteProtection()->ValidatePassword(u"MyPassword"));
    ASSERT_FALSE(doc->get_WriteProtection()->ValidatePassword(u"wrongpassword"));
}

namespace gtest_test
{

TEST_F(ExDocument, WriteProtection)
{
    s_instance->WriteProtection();
}

} // namespace gtest_test

void ExDocument::RemovePersonalInformation(bool saveWithoutPersonalInfo)
{
    //ExStart
    //ExFor:Document.RemovePersonalInformation
    //ExSummary:Shows how to enable the removal of personal information during a manual save.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert some content with personal information.
    doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
    doc->get_BuiltInDocumentProperties()->set_Company(u"Placeholder Inc.");
    
    doc->StartTrackRevisions(doc->get_BuiltInDocumentProperties()->get_Author(), System::DateTime::get_Now());
    builder->Write(u"Hello world!");
    doc->StopTrackRevisions();
    
    // This flag is equivalent to File -> Options -> Trust Center -> Trust Center Settings... ->
    // Privacy Options -> "Remove personal information from file properties on save" in Microsoft Word.
    doc->set_RemovePersonalInformation(saveWithoutPersonalInfo);
    
    // This option will not take effect during a save operation made using Aspose.Words.
    // Personal data will be removed from our document with the flag set when we save it manually using Microsoft Word.
    doc->Save(get_ArtifactsDir() + u"Document.RemovePersonalInformation.docx");
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.RemovePersonalInformation.docx");
    
    ASPOSE_ASSERT_EQ(saveWithoutPersonalInfo, doc->get_RemovePersonalInformation());
    ASSERT_EQ(u"John Doe", doc->get_BuiltInDocumentProperties()->get_Author());
    ASSERT_EQ(u"Placeholder Inc.", doc->get_BuiltInDocumentProperties()->get_Company());
    ASSERT_EQ(u"John Doe", doc->get_Revisions()->idx_get(0)->get_Author());
    //ExEnd
}

namespace gtest_test
{

using ExDocument_RemovePersonalInformation_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocument::RemovePersonalInformation)>::type;

struct ExDocument_RemovePersonalInformation : public ExDocument, public Aspose::Words::ApiExamples::ExDocument, public ::testing::WithParamInterface<ExDocument_RemovePersonalInformation_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocument_RemovePersonalInformation, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RemovePersonalInformation(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_RemovePersonalInformation, ::testing::ValuesIn(ExDocument_RemovePersonalInformation::TestCases()));

} // namespace gtest_test

void ExDocument::ShowComments()
{
    //ExStart
    //ExFor:LayoutOptions.CommentDisplayMode
    //ExFor:CommentDisplayMode
    //ExSummary:Shows how to show comments when saving a document to a rendered format.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Hello world!");
    
    auto comment = System::MakeObject<Aspose::Words::Comment>(doc, u"John Doe", u"J.D.", System::DateTime::get_Now());
    comment->SetText(u"My comment.");
    builder->get_CurrentParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Comment>>(comment);
    
    // ShowInAnnotations is only available in Pdf1.7 and Pdf1.5 formats.
    // In other formats, it will work similarly to Hide.
    doc->get_LayoutOptions()->set_CommentDisplayMode(Aspose::Words::Layout::CommentDisplayMode::ShowInAnnotations);
    
    doc->Save(get_ArtifactsDir() + u"Document.ShowCommentsInAnnotations.pdf");
    
    // Note that it's required to rebuild the document page layout (via Document.UpdatePageLayout() method)
    // after changing the Document.LayoutOptions values.
    doc->get_LayoutOptions()->set_CommentDisplayMode(Aspose::Words::Layout::CommentDisplayMode::ShowInBalloons);
    doc->UpdatePageLayout();
    
    doc->Save(get_ArtifactsDir() + u"Document.ShowCommentsInBalloons.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, ShowComments)
{
    s_instance->ShowComments();
}

} // namespace gtest_test

void ExDocument::CopyTemplateStylesViaDocument()
{
    //ExStart
    //ExFor:Document.CopyStylesFromTemplate(Document)
    //ExSummary:Shows how to copies styles from the template to a document via Document.
    auto template_ = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    auto target = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    ASSERT_EQ(18, template_->get_Styles()->get_Count());
    //ExSkip
    ASSERT_EQ(12, target->get_Styles()->get_Count());
    //ExSkip
    
    target->CopyStylesFromTemplate(template_);
    ASSERT_EQ(22, target->get_Styles()->get_Count());
    //ExSkip
    
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, CopyTemplateStylesViaDocument)
{
    s_instance->CopyTemplateStylesViaDocument();
}

} // namespace gtest_test

void ExDocument::CopyTemplateStylesViaDocumentNew()
{
    //ExStart
    //ExFor:Document.CopyStylesFromTemplate(Document)
    //ExFor:Document.CopyStylesFromTemplate(String)
    //ExSummary:Shows how to copy styles from one document to another.
    // Create a document, and then add styles that we will copy to another document.
    auto template_ = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Style> style = template_->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"TemplateStyle1");
    style->get_Font()->set_Name(u"Times New Roman");
    style->get_Font()->set_Color(System::Drawing::Color::get_Navy());
    
    style = template_->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"TemplateStyle2");
    style->get_Font()->set_Name(u"Arial");
    style->get_Font()->set_Color(System::Drawing::Color::get_DeepSkyBlue());
    
    style = template_->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"TemplateStyle3");
    style->get_Font()->set_Name(u"Courier New");
    style->get_Font()->set_Color(System::Drawing::Color::get_RoyalBlue());
    
    ASSERT_EQ(7, template_->get_Styles()->get_Count());
    
    // Create a document which we will copy the styles to.
    auto target = System::MakeObject<Aspose::Words::Document>();
    
    // Create a style with the same name as a style from the template document and add it to the target document.
    style = target->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"TemplateStyle3");
    style->get_Font()->set_Name(u"Calibri");
    style->get_Font()->set_Color(System::Drawing::Color::get_Orange());
    
    ASSERT_EQ(5, target->get_Styles()->get_Count());
    
    // There are two ways of calling the method to copy all the styles from one document to another.
    // 1 -  Passing the template document object:
    target->CopyStylesFromTemplate(template_);
    
    // Copying styles adds all styles from the template document to the target
    // and overwrites existing styles with the same name.
    ASSERT_EQ(7, target->get_Styles()->get_Count());
    
    ASSERT_EQ(u"Courier New", target->get_Styles()->idx_get(u"TemplateStyle3")->get_Font()->get_Name());
    ASSERT_EQ(System::Drawing::Color::get_RoyalBlue().ToArgb(), target->get_Styles()->idx_get(u"TemplateStyle3")->get_Font()->get_Color().ToArgb());
    
    // 2 -  Passing the local system filename of a template document:
    target->CopyStylesFromTemplate(get_MyDir() + u"Rendering.docx");
    
    ASSERT_EQ(21, target->get_Styles()->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, CopyTemplateStylesViaDocumentNew)
{
    s_instance->CopyTemplateStylesViaDocumentNew();
}

} // namespace gtest_test

void ExDocument::ReadMacrosFromExistingDocument()
{
    //ExStart
    //ExFor:Document.VbaProject
    //ExFor:VbaModuleCollection
    //ExFor:VbaModuleCollection.Count
    //ExFor:VbaModuleCollection.Item(System.Int32)
    //ExFor:VbaModuleCollection.Item(System.String)
    //ExFor:VbaModuleCollection.Remove
    //ExFor:VbaModule
    //ExFor:VbaModule.Name
    //ExFor:VbaModule.SourceCode
    //ExFor:VbaProject
    //ExFor:VbaProject.Name
    //ExFor:VbaProject.Modules
    //ExFor:VbaProject.CodePage
    //ExFor:VbaProject.IsSigned
    //ExSummary:Shows how to access a document's VBA project information.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"VBA project.docm");
    
    // A VBA project contains a collection of VBA modules.
    System::SharedPtr<Aspose::Words::Vba::VbaProject> vbaProject = doc->get_VbaProject();
    ASSERT_TRUE(vbaProject->get_IsSigned());
    //ExSkip
    std::cout << (vbaProject->get_IsSigned() ? System::String::Format(u"Project name: {0} signed; Project code page: {1}; Modules count: {2}\n", vbaProject->get_Name(), vbaProject->get_CodePage(), vbaProject->get_Modules()->LINQ_Count()) : System::String::Format(u"Project name: {0} not signed; Project code page: {1}; Modules count: {2}\n", vbaProject->get_Name(), vbaProject->get_CodePage(), vbaProject->get_Modules()->LINQ_Count())) << std::endl;
    
    System::SharedPtr<Aspose::Words::Vba::VbaModuleCollection> vbaModules = doc->get_VbaProject()->get_Modules();
    
    ASSERT_EQ(vbaModules->LINQ_Count(), 3);
    
    for (auto&& module_ : vbaModules)
    {
        std::cout << System::String::Format(u"Module name: {0};\nModule code:\n{1}\n", module_->get_Name(), module_->get_SourceCode()) << std::endl;
    }
    
    // Set new source code for VBA module. You can access VBA modules in the collection either by index or by name.
    vbaModules->idx_get(0)->set_SourceCode(u"Your VBA code...");
    vbaModules->idx_get(u"Module1")->set_SourceCode(u"Your VBA code...");
    
    // Remove a module from the collection.
    vbaModules->Remove(vbaModules->idx_get(2));
    //ExEnd
    
    ASSERT_EQ(u"AsposeVBAtest", vbaProject->get_Name());
    ASSERT_EQ(2, vbaProject->get_Modules()->LINQ_Count());
    ASSERT_EQ(1251, vbaProject->get_CodePage());
    ASSERT_FALSE(vbaProject->get_IsSigned());
    
    ASSERT_EQ(u"ThisDocument", vbaModules->idx_get(0)->get_Name());
    ASSERT_EQ(u"Your VBA code...", vbaModules->idx_get(0)->get_SourceCode());
    
    ASSERT_EQ(u"Module1", vbaModules->idx_get(1)->get_Name());
    ASSERT_EQ(u"Your VBA code...", vbaModules->idx_get(1)->get_SourceCode());
}

namespace gtest_test
{

TEST_F(ExDocument, ReadMacrosFromExistingDocument)
{
    s_instance->ReadMacrosFromExistingDocument();
}

} // namespace gtest_test

void ExDocument::SaveOutputParameters()
{
    //ExStart
    //ExFor:SaveOutputParameters
    //ExFor:SaveOutputParameters.ContentType
    //ExSummary:Shows how to access output parameters of a document's save operation.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    // After we save a document, we can access the Internet Media Type (MIME type) of the newly created output document.
    System::SharedPtr<Aspose::Words::Saving::SaveOutputParameters> parameters = doc->Save(get_ArtifactsDir() + u"Document.SaveOutputParameters.doc");
    
    ASSERT_EQ(u"application/msword", parameters->get_ContentType());
    
    // This property changes depending on the save format.
    parameters = doc->Save(get_ArtifactsDir() + u"Document.SaveOutputParameters.pdf");
    
    ASSERT_EQ(u"application/pdf", parameters->get_ContentType());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, SaveOutputParameters)
{
    s_instance->SaveOutputParameters();
}

} // namespace gtest_test

void ExDocument::SubDocument()
{
    //ExStart
    //ExFor:SubDocument
    //ExFor:SubDocument.NodeType
    //ExSummary:Shows how to access a master document's subdocument.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Master document.docx");
    
    System::SharedPtr<Aspose::Words::NodeCollection> subDocuments = doc->GetChildNodes(Aspose::Words::NodeType::SubDocument, true);
    ASSERT_EQ(1, subDocuments->get_Count());
    //ExSkip
    
    // This node serves as a reference to an external document, and its contents cannot be accessed.
    auto subDocument = System::ExplicitCast<Aspose::Words::SubDocument>(subDocuments->idx_get(0));
    
    ASSERT_FALSE(subDocument->get_IsComposite());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, SubDocument)
{
    s_instance->SubDocument();
}

} // namespace gtest_test

void ExDocument::CreateWebExtension()
{
    //ExStart
    //ExFor:BaseWebExtensionCollection`1.Add(`0)
    //ExFor:BaseWebExtensionCollection`1.Clear
    //ExFor:Document.WebExtensionTaskPanes
    //ExFor:TaskPane
    //ExFor:TaskPane.DockState
    //ExFor:TaskPane.IsVisible
    //ExFor:TaskPane.Width
    //ExFor:TaskPane.IsLocked
    //ExFor:TaskPane.WebExtension
    //ExFor:TaskPane.Row
    //ExFor:WebExtension
    //ExFor:WebExtension.Id
    //ExFor:WebExtension.AlternateReferences
    //ExFor:WebExtension.Reference
    //ExFor:WebExtension.Properties
    //ExFor:WebExtension.Bindings
    //ExFor:WebExtension.IsFrozen
    //ExFor:WebExtensionReference
    //ExFor:WebExtensionReference.Id
    //ExFor:WebExtensionReference.Version
    //ExFor:WebExtensionReference.StoreType
    //ExFor:WebExtensionReference.Store
    //ExFor:WebExtensionPropertyCollection
    //ExFor:WebExtensionBindingCollection
    //ExFor:WebExtensionProperty.#ctor(String, String)
    //ExFor:WebExtensionProperty.Name
    //ExFor:WebExtensionProperty.Value
    //ExFor:WebExtensionBinding.#ctor(String, WebExtensionBindingType, String)
    //ExFor:WebExtensionStoreType
    //ExFor:WebExtensionBindingType
    //ExFor:TaskPaneDockState
    //ExFor:TaskPaneCollection
    //ExFor:WebExtensionBinding.Id
    //ExFor:WebExtensionBinding.AppRef
    //ExFor:WebExtensionBinding.BindingType
    //ExSummary:Shows how to add a web extension to a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create task pane with "MyScript" add-in, which will be used by the document,
    // then set its default location.
    auto myScriptTaskPane = System::MakeObject<Aspose::Words::WebExtensions::TaskPane>();
    doc->get_WebExtensionTaskPanes()->Add(myScriptTaskPane);
    myScriptTaskPane->set_DockState(Aspose::Words::WebExtensions::TaskPaneDockState::Right);
    myScriptTaskPane->set_IsVisible(true);
    myScriptTaskPane->set_Width(300);
    myScriptTaskPane->set_IsLocked(true);
    
    // If there are multiple task panes in the same docking location, we can set this index to arrange them.
    myScriptTaskPane->set_Row(1);
    
    // Create an add-in called "MyScript Math Sample", which the task pane will display within.
    System::SharedPtr<Aspose::Words::WebExtensions::WebExtension> webExtension = myScriptTaskPane->get_WebExtension();
    
    // Set application store reference parameters for our add-in, such as the ID.
    webExtension->get_Reference()->set_Id(u"WA104380646");
    webExtension->get_Reference()->set_Version(u"1.0.0.0");
    webExtension->get_Reference()->set_StoreType(Aspose::Words::WebExtensions::WebExtensionStoreType::OMEX);
    webExtension->get_Reference()->set_Store(System::Globalization::CultureInfo::get_CurrentCulture()->get_Name());
    webExtension->get_Properties()->Add(System::MakeObject<Aspose::Words::WebExtensions::WebExtensionProperty>(u"MyScript", u"MyScript Math Sample"));
    webExtension->get_Bindings()->Add(System::MakeObject<Aspose::Words::WebExtensions::WebExtensionBinding>(u"MyScript", Aspose::Words::WebExtensions::WebExtensionBindingType::Text, u"104380646"));
    
    // Allow the user to interact with the add-in.
    webExtension->set_IsFrozen(false);
    
    // We can access the web extension in Microsoft Word via Developer -> Add-ins.
    doc->Save(get_ArtifactsDir() + u"Document.WebExtension.docx");
    
    // Remove all web extension task panes at once like this.
    doc->get_WebExtensionTaskPanes()->Clear();
    
    ASSERT_EQ(0, doc->get_WebExtensionTaskPanes()->get_Count());
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.WebExtension.docx");
    
    myScriptTaskPane = doc->get_WebExtensionTaskPanes()->idx_get(0);
    ASSERT_EQ(Aspose::Words::WebExtensions::TaskPaneDockState::Right, myScriptTaskPane->get_DockState());
    ASSERT_TRUE(myScriptTaskPane->get_IsVisible());
    ASPOSE_ASSERT_EQ(300.0, myScriptTaskPane->get_Width());
    ASSERT_TRUE(myScriptTaskPane->get_IsLocked());
    ASSERT_EQ(1, myScriptTaskPane->get_Row());
    
    webExtension = myScriptTaskPane->get_WebExtension();
    ASSERT_EQ(System::String::Empty, webExtension->get_Id());
    
    ASSERT_EQ(u"WA104380646", webExtension->get_Reference()->get_Id());
    ASSERT_EQ(u"1.0.0.0", webExtension->get_Reference()->get_Version());
    ASSERT_EQ(Aspose::Words::WebExtensions::WebExtensionStoreType::OMEX, webExtension->get_Reference()->get_StoreType());
    ASSERT_EQ(System::Globalization::CultureInfo::get_CurrentCulture()->get_Name(), webExtension->get_Reference()->get_Store());
    ASSERT_EQ(0, webExtension->get_AlternateReferences()->get_Count());
    
    ASSERT_EQ(u"MyScript", webExtension->get_Properties()->idx_get(0)->get_Name());
    ASSERT_EQ(u"MyScript Math Sample", webExtension->get_Properties()->idx_get(0)->get_Value());
    
    ASSERT_EQ(u"MyScript", webExtension->get_Bindings()->idx_get(0)->get_Id());
    ASSERT_EQ(Aspose::Words::WebExtensions::WebExtensionBindingType::Text, webExtension->get_Bindings()->idx_get(0)->get_BindingType());
    ASSERT_EQ(u"104380646", webExtension->get_Bindings()->idx_get(0)->get_AppRef());
    
    ASSERT_FALSE(webExtension->get_IsFrozen());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, CreateWebExtension)
{
    s_instance->CreateWebExtension();
}

} // namespace gtest_test

void ExDocument::GetWebExtensionInfo()
{
    //ExStart
    //ExFor:BaseWebExtensionCollection`1
    //ExFor:BaseWebExtensionCollection`1.GetEnumerator
    //ExFor:BaseWebExtensionCollection`1.Remove(Int32)
    //ExFor:BaseWebExtensionCollection`1.Count
    //ExFor:BaseWebExtensionCollection`1.Item(Int32)
    //ExSummary:Shows how to work with a document's collection of web extensions.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Web extension.docx");
    
    ASSERT_EQ(1, doc->get_WebExtensionTaskPanes()->get_Count());
    
    // Print all properties of the document's web extension.
    System::SharedPtr<Aspose::Words::WebExtensions::WebExtensionPropertyCollection> webExtensionPropertyCollection = doc->get_WebExtensionTaskPanes()->idx_get(0)->get_WebExtension()->get_Properties();
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::WebExtensions::WebExtensionProperty>>> enumerator = webExtensionPropertyCollection->GetEnumerator();
        while (enumerator->MoveNext())
        {
            System::SharedPtr<Aspose::Words::WebExtensions::WebExtensionProperty> webExtensionProperty = enumerator->get_Current();
            std::cout << System::String::Format(u"Binding name: {0}; Binding value: {1}", webExtensionProperty->get_Name(), webExtensionProperty->get_Value()) << std::endl;
        }
    }
    
    // Remove the web extension.
    doc->get_WebExtensionTaskPanes()->Remove(0);
    
    ASSERT_EQ(0, doc->get_WebExtensionTaskPanes()->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, GetWebExtensionInfo)
{
    s_instance->GetWebExtensionInfo();
}

} // namespace gtest_test

void ExDocument::EpubCover()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    // When saving to .epub, some Microsoft Word document properties convert to .epub metadata.
    doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
    doc->get_BuiltInDocumentProperties()->set_Title(u"My Book Title");
    
    // The thumbnail we specify here can become the cover image.
    System::ArrayPtr<uint8_t> image = System::IO::File::ReadAllBytes(get_ImageDir() + u"Transparent background logo.png");
    doc->get_BuiltInDocumentProperties()->set_Thumbnail(image);
    
    doc->Save(get_ArtifactsDir() + u"Document.EpubCover.epub");
}

namespace gtest_test
{

TEST_F(ExDocument, EpubCover)
{
    s_instance->EpubCover();
}

} // namespace gtest_test

void ExDocument::TextWatermark()
{
    //ExStart
    //ExFor:Document.Watermark
    //ExFor:Watermark
    //ExFor:Watermark.SetText(String)
    //ExFor:Watermark.SetText(String, TextWatermarkOptions)
    //ExFor:Watermark.Remove
    //ExFor:TextWatermarkOptions
    //ExFor:TextWatermarkOptions.FontFamily
    //ExFor:TextWatermarkOptions.FontSize
    //ExFor:TextWatermarkOptions.Color
    //ExFor:TextWatermarkOptions.Layout
    //ExFor:TextWatermarkOptions.IsSemitrasparent
    //ExFor:WatermarkLayout
    //ExFor:WatermarkType
    //ExFor:Watermark.Type
    //ExSummary:Shows how to create a text watermark.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Add a plain text watermark.
    doc->get_Watermark()->SetText(u"Aspose Watermark");
    
    // If we wish to edit the text formatting using it as a watermark,
    // we can do so by passing a TextWatermarkOptions object when creating the watermark.
    auto textWatermarkOptions = System::MakeObject<Aspose::Words::TextWatermarkOptions>();
    textWatermarkOptions->set_FontFamily(u"Arial");
    textWatermarkOptions->set_FontSize(36.0f);
    textWatermarkOptions->set_Color(System::Drawing::Color::get_Black());
    textWatermarkOptions->set_Layout(Aspose::Words::WatermarkLayout::Diagonal);
    textWatermarkOptions->set_IsSemitrasparent(false);
    
    doc->get_Watermark()->SetText(u"Aspose Watermark", textWatermarkOptions);
    
    doc->Save(get_ArtifactsDir() + u"Document.TextWatermark.docx");
    
    // We can remove a watermark from a document like this.
    if (doc->get_Watermark()->get_Type() == Aspose::Words::WatermarkType::Text)
    {
        doc->get_Watermark()->Remove();
    }
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.TextWatermark.docx");
    
    ASSERT_EQ(Aspose::Words::WatermarkType::Text, doc->get_Watermark()->get_Type());
}

namespace gtest_test
{

TEST_F(ExDocument, TextWatermark)
{
    s_instance->TextWatermark();
}

} // namespace gtest_test

void ExDocument::ImageWatermark()
{
    //ExStart
    //ExFor:Watermark.SetImage(Image)
    //ExFor:Watermark.SetImage(Image, ImageWatermarkOptions)
    //ExFor:Watermark.SetImage(String, ImageWatermarkOptions)
    //ExFor:ImageWatermarkOptions
    //ExFor:ImageWatermarkOptions.Scale
    //ExFor:ImageWatermarkOptions.IsWashout
    //ExSummary:Shows how to create a watermark from an image in the local file system.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Modify the image watermark's appearance with an ImageWatermarkOptions object,
    // then pass it while creating a watermark from an image file.
    auto imageWatermarkOptions = System::MakeObject<Aspose::Words::ImageWatermarkOptions>();
    imageWatermarkOptions->set_Scale(5);
    imageWatermarkOptions->set_IsWashout(false);
    
    // We have a different options to insert image.
    // Use on of the following methods to add image watermark.
    doc->get_Watermark()->SetImage(System::Drawing::Image::FromFile(get_ImageDir() + u"Logo.jpg"));
    
    doc->get_Watermark()->SetImage(System::Drawing::Image::FromFile(get_ImageDir() + u"Logo.jpg"), imageWatermarkOptions);
    
    doc->get_Watermark()->SetImage(get_ImageDir() + u"Logo.jpg", imageWatermarkOptions);
    
    
    doc->Save(get_ArtifactsDir() + u"Document.ImageWatermark.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.ImageWatermark.docx");
    ASSERT_EQ(Aspose::Words::WatermarkType::Image, doc->get_Watermark()->get_Type());
}

namespace gtest_test
{

TEST_F(ExDocument, ImageWatermark)
{
    s_instance->ImageWatermark();
}

} // namespace gtest_test

void ExDocument::ImageWatermarkStream()
{
    //ExStart:ImageWatermarkStream
    //GistId:12a3a3cfe30f3145220db88428a9f814
    //ExFor:Watermark.SetImage(Stream, ImageWatermarkOptions)
    //ExSummary:Shows how to create a watermark from an image stream.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Modify the image watermark's appearance with an ImageWatermarkOptions object,
    // then pass it while creating a watermark from an image file.
    auto imageWatermarkOptions = System::MakeObject<Aspose::Words::ImageWatermarkOptions>();
    imageWatermarkOptions->set_Scale(5);
    
    {
        auto imageStream = System::MakeObject<System::IO::FileStream>(get_ImageDir() + u"Logo.jpg", System::IO::FileMode::Open, System::IO::FileAccess::Read);
        doc->get_Watermark()->SetImage(imageStream, imageWatermarkOptions);
    }
    
    doc->Save(get_ArtifactsDir() + u"Document.ImageWatermarkStream.docx");
    //ExEnd:ImageWatermarkStream
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.ImageWatermarkStream.docx");
    ASSERT_EQ(Aspose::Words::WatermarkType::Image, doc->get_Watermark()->get_Type());
}

namespace gtest_test
{

TEST_F(ExDocument, ImageWatermarkStream)
{
    s_instance->ImageWatermarkStream();
}

} // namespace gtest_test

void ExDocument::SpellingAndGrammarErrors(bool showErrors)
{
    //ExStart
    //ExFor:Document.ShowGrammaticalErrors
    //ExFor:Document.ShowSpellingErrors
    //ExSummary:Shows how to show/hide errors in the document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert two sentences with mistakes that would be picked up
    // by the spelling and grammar checkers in Microsoft Word.
    builder->Writeln(u"There is a speling error in this sentence.");
    builder->Writeln(u"Their is a grammatical error in this sentence.");
    
    // If these options are enabled, then spelling errors will be underlined
    // in the output document by a jagged red line, and a double blue line will highlight grammatical mistakes.
    doc->set_ShowGrammaticalErrors(showErrors);
    doc->set_ShowSpellingErrors(showErrors);
    
    doc->Save(get_ArtifactsDir() + u"Document.SpellingAndGrammarErrors.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.SpellingAndGrammarErrors.docx");
    
    ASPOSE_ASSERT_EQ(showErrors, doc->get_ShowGrammaticalErrors());
    ASPOSE_ASSERT_EQ(showErrors, doc->get_ShowSpellingErrors());
}

namespace gtest_test
{

using ExDocument_SpellingAndGrammarErrors_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocument::SpellingAndGrammarErrors)>::type;

struct ExDocument_SpellingAndGrammarErrors : public ExDocument, public Aspose::Words::ApiExamples::ExDocument, public ::testing::WithParamInterface<ExDocument_SpellingAndGrammarErrors_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocument_SpellingAndGrammarErrors, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SpellingAndGrammarErrors(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_SpellingAndGrammarErrors, ::testing::ValuesIn(ExDocument_SpellingAndGrammarErrors::TestCases()));

} // namespace gtest_test

void ExDocument::IgnorePrinterMetrics()
{
    //ExStart
    //ExFor:LayoutOptions.IgnorePrinterMetrics
    //ExSummary:Shows how to ignore 'Use printer metrics to lay out document' option.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    doc->get_LayoutOptions()->set_IgnorePrinterMetrics(false);
    
    doc->Save(get_ArtifactsDir() + u"Document.IgnorePrinterMetrics.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, IgnorePrinterMetrics)
{
    s_instance->IgnorePrinterMetrics();
}

} // namespace gtest_test

void ExDocument::ExtractPages()
{
    //ExStart
    //ExFor:Document.ExtractPages(int, int)
    //ExSummary:Shows how to get specified range of pages from the document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Layout entities.docx");
    
    doc = doc->ExtractPages(0, 2);
    
    doc->Save(get_ArtifactsDir() + u"Document.ExtractPages.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Document.ExtractPages.docx");
    ASSERT_EQ(doc->get_PageCount(), 2);
}

namespace gtest_test
{

TEST_F(ExDocument, ExtractPages)
{
    s_instance->ExtractPages();
}

} // namespace gtest_test

void ExDocument::SpellingOrGrammar(bool checkSpellingGrammar)
{
    //ExStart
    //ExFor:Document.SpellingChecked
    //ExFor:Document.GrammarChecked
    //ExSummary:Shows how to set spelling or grammar verifying.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // The string with spelling errors.
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->Add(System::MakeObject<Aspose::Words::Run>(doc, u"The speeling in this documentz is all broked."));
    
    // Spelling/Grammar check start if we set properties to false.
    // We can see all errors in Microsoft Word via Review -> Spelling & Grammar.
    // Note that Microsoft Word does not start grammar/spell check automatically for DOC and RTF document format.
    doc->set_SpellingChecked(checkSpellingGrammar);
    doc->set_GrammarChecked(checkSpellingGrammar);
    
    doc->Save(get_ArtifactsDir() + u"Document.SpellingOrGrammar.docx");
    //ExEnd
}

namespace gtest_test
{

using ExDocument_SpellingOrGrammar_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocument::SpellingOrGrammar)>::type;

struct ExDocument_SpellingOrGrammar : public ExDocument, public Aspose::Words::ApiExamples::ExDocument, public ::testing::WithParamInterface<ExDocument_SpellingOrGrammar_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExDocument_SpellingOrGrammar, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SpellingOrGrammar(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_SpellingOrGrammar, ::testing::ValuesIn(ExDocument_SpellingOrGrammar::TestCases()));

} // namespace gtest_test

void ExDocument::AllowEmbeddingPostScriptFonts()
{
    //ExStart
    //ExFor:SaveOptions.AllowEmbeddingPostScriptFonts
    //ExSummary:Shows how to save the document with PostScript font.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"PostScriptFont");
    builder->Writeln(u"Some text with PostScript font.");
    
    // Load the font with PostScript to use in the document.
    auto otf = System::MakeObject<Aspose::Words::Fonts::MemoryFontSource>(System::IO::File::ReadAllBytes(get_FontsDir() + u"AllegroOpen.otf"));
    doc->set_FontSettings(System::MakeObject<Aspose::Words::Fonts::FontSettings>());
    doc->get_FontSettings()->SetFontsSources(System::MakeArray<System::SharedPtr<Aspose::Words::Fonts::FontSourceBase>>({otf}));
    
    // Embed TrueType fonts.
    doc->get_FontInfos()->set_EmbedTrueTypeFonts(true);
    
    // Allow embedding PostScript fonts while embedding TrueType fonts.
    // Microsoft Word does not embed PostScript fonts, but can open documents with embedded fonts of this type.
    System::SharedPtr<Aspose::Words::Saving::SaveOptions> saveOptions = Aspose::Words::Saving::SaveOptions::CreateSaveOptions(Aspose::Words::SaveFormat::Docx);
    saveOptions->set_AllowEmbeddingPostScriptFonts(true);
    
    doc->Save(get_ArtifactsDir() + u"Document.AllowEmbeddingPostScriptFonts.docx", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, AllowEmbeddingPostScriptFonts)
{
    s_instance->AllowEmbeddingPostScriptFonts();
}

} // namespace gtest_test

void ExDocument::Frameset()
{
    //ExStart
    //ExFor:Document.Frameset
    //ExFor:Frameset
    //ExFor:Frameset.FrameDefaultUrl
    //ExFor:Frameset.IsFrameLinkToFile
    //ExFor:Frameset.ChildFramesets
    //ExFor:FramesetCollection
    //ExFor:FramesetCollection.Count
    //ExFor:FramesetCollection.Item(Int32)
    //ExSummary:Shows how to access frames on-page.
    // Document contains several frames with links to other documents.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Frameset.docx");
    
    ASSERT_EQ(3, doc->get_Frameset()->get_ChildFramesets()->get_Count());
    // We can check the default URL (a web page URL or local document) or if the frame is an external resource.
    ASSERT_EQ(u"https://file-examples-com.github.io/uploads/2017/02/file-sample_100kB.docx", doc->get_Frameset()->get_ChildFramesets()->idx_get(0)->get_ChildFramesets()->idx_get(0)->get_FrameDefaultUrl());
    ASSERT_TRUE(doc->get_Frameset()->get_ChildFramesets()->idx_get(0)->get_ChildFramesets()->idx_get(0)->get_IsFrameLinkToFile());
    
    ASSERT_EQ(u"Document.docx", doc->get_Frameset()->get_ChildFramesets()->idx_get(1)->get_FrameDefaultUrl());
    ASSERT_FALSE(doc->get_Frameset()->get_ChildFramesets()->idx_get(1)->get_IsFrameLinkToFile());
    
    // Change properties for one of our frames.
    doc->get_Frameset()->get_ChildFramesets()->idx_get(0)->get_ChildFramesets()->idx_get(0)->set_FrameDefaultUrl(u"https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx");
    doc->get_Frameset()->get_ChildFramesets()->idx_get(0)->get_ChildFramesets()->idx_get(0)->set_IsFrameLinkToFile(false);
    //ExEnd
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    
    ASSERT_EQ(u"https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx", doc->get_Frameset()->get_ChildFramesets()->idx_get(0)->get_ChildFramesets()->idx_get(0)->get_FrameDefaultUrl());
    ASSERT_FALSE(doc->get_Frameset()->get_ChildFramesets()->idx_get(0)->get_ChildFramesets()->idx_get(0)->get_IsFrameLinkToFile());
}

namespace gtest_test
{

TEST_F(ExDocument, Frameset)
{
    s_instance->Frameset();
}

} // namespace gtest_test

void ExDocument::OpenAzw()
{
    System::SharedPtr<Aspose::Words::FileFormatInfo> info = Aspose::Words::FileFormatUtil::DetectFileFormat(get_MyDir() + u"Azw3 document.azw3");
    ASSERT_EQ(info->get_LoadFormat(), Aspose::Words::LoadFormat::Azw3);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Azw3 document.azw3");
    ASSERT_TRUE(doc->GetText().Contains(u"Hachette Book Group USA"));
}

namespace gtest_test
{

TEST_F(ExDocument, OpenAzw)
{
    s_instance->OpenAzw();
}

} // namespace gtest_test

void ExDocument::OpenEpub()
{
    System::SharedPtr<Aspose::Words::FileFormatInfo> info = Aspose::Words::FileFormatUtil::DetectFileFormat(get_MyDir() + u"Epub document.epub");
    ASSERT_EQ(info->get_LoadFormat(), Aspose::Words::LoadFormat::Epub);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Epub document.epub");
    ASSERT_TRUE(doc->GetText().Contains(u"Down the Rabbit-Hole"));
}

namespace gtest_test
{

TEST_F(ExDocument, OpenEpub)
{
    s_instance->OpenEpub();
}

} // namespace gtest_test

void ExDocument::OpenXml()
{
    System::SharedPtr<Aspose::Words::FileFormatInfo> info = Aspose::Words::FileFormatUtil::DetectFileFormat(get_MyDir() + u"Mail merge data - Customers.xml");
    ASSERT_EQ(info->get_LoadFormat(), Aspose::Words::LoadFormat::Xml);
    
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Mail merge data - Purchase order.xml");
    ASSERT_TRUE(doc->GetText().Contains(u"Ellen Adams\r123 Maple Street"));
}

namespace gtest_test
{

TEST_F(ExDocument, OpenXml)
{
    s_instance->OpenXml();
}

} // namespace gtest_test

void ExDocument::MoveToStructuredDocumentTag()
{
    //ExStart
    //ExFor:DocumentBuilder.MoveToStructuredDocumentTag(int, int)
    //ExFor:DocumentBuilder.MoveToStructuredDocumentTag(StructuredDocumentTag, int)
    //ExFor:DocumentBuilder.IsAtEndOfStructuredDocumentTag
    //ExFor:DocumentBuilder.CurrentStructuredDocumentTag
    //ExSummary:Shows how to move cursor of DocumentBuilder inside a structured document tag.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Structured document tags.docx");
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // There is a several ways to move the cursor:
    // 1 -  Move to the first character of structured document tag by index.
    builder->MoveToStructuredDocumentTag(1, 1);
    
    // 2 -  Move to the first character of structured document tag by object.
    auto tag = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTag, 2, true));
    builder->MoveToStructuredDocumentTag(tag, 1);
    builder->Write(u" New text.");
    
    ASSERT_EQ(u"R New text.ichText", tag->GetText().Trim());
    
    // 3 -  Move to the end of the second structured document tag.
    builder->MoveToStructuredDocumentTag(1, -1);
    ASSERT_TRUE(builder->get_IsAtEndOfStructuredDocumentTag());
    
    // Get currently selected structured document tag.
    builder->get_CurrentStructuredDocumentTag()->set_Color(System::Drawing::Color::get_Green());
    
    doc->Save(get_ArtifactsDir() + u"Document.MoveToStructuredDocumentTag.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, MoveToStructuredDocumentTag)
{
    s_instance->MoveToStructuredDocumentTag();
}

} // namespace gtest_test

void ExDocument::IncludeTextboxesFootnotesEndnotesInStat()
{
    //ExStart
    //ExFor:Document.IncludeTextboxesFootnotesEndnotesInStat
    //ExSummary: Shows how to include or exclude textboxes, footnotes and endnotes from word count statistics.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Lorem ipsum");
    builder->InsertFootnote(Aspose::Words::Notes::FootnoteType::Footnote, u"sit amet");
    
    // By default option is set to 'false'.
    doc->UpdateWordCount();
    // Words count without textboxes, footnotes and endnotes.
    ASSERT_EQ(2, doc->get_BuiltInDocumentProperties()->get_Words());
    
    doc->set_IncludeTextboxesFootnotesEndnotesInStat(true);
    doc->UpdateWordCount();
    // Words count with textboxes, footnotes and endnotes.
    ASSERT_EQ(4, doc->get_BuiltInDocumentProperties()->get_Words());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, IncludeTextboxesFootnotesEndnotesInStat)
{
    s_instance->IncludeTextboxesFootnotesEndnotesInStat();
}

} // namespace gtest_test

void ExDocument::SetJustificationMode()
{
    //ExStart
    //ExFor:Document.JustificationMode
    //ExFor:JustificationMode
    //ExSummary:Shows how to manage character spacing control.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    Aspose::Words::Settings::JustificationMode justificationMode = doc->get_JustificationMode();
    if (justificationMode == Aspose::Words::Settings::JustificationMode::Expand)
    {
        doc->set_JustificationMode(Aspose::Words::Settings::JustificationMode::Compress);
    }
    
    doc->Save(get_ArtifactsDir() + u"Document.SetJustificationMode.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, SetJustificationMode)
{
    s_instance->SetJustificationMode();
}

} // namespace gtest_test

void ExDocument::PageIsInColor()
{
    //ExStart
    //ExFor:PageInfo.Colored
    //ExFor:Document.GetPageInfo(Int32)
    //ExSummary:Shows how to check whether the page is in color or not.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    // Check that the first page of the document is not colored.
    ASSERT_FALSE(doc->GetPageInfo(0)->get_Colored());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, PageIsInColor)
{
    s_instance->PageIsInColor();
}

} // namespace gtest_test

void ExDocument::InsertDocumentInline()
{
    //ExStart:InsertDocumentInline
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:DocumentBuilder.InsertDocumentInline(Document, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to insert a document inline at the cursor position.
    auto srcDoc = System::MakeObject<Aspose::Words::DocumentBuilder>();
    srcDoc->Write(u"[src content]");
    
    // Create destination document.
    auto dstDoc = System::MakeObject<Aspose::Words::DocumentBuilder>();
    dstDoc->Write(u"Before ");
    dstDoc->InsertNode(System::MakeObject<Aspose::Words::BookmarkStart>(dstDoc->get_Document(), u"src_place"));
    dstDoc->InsertNode(System::MakeObject<Aspose::Words::BookmarkEnd>(dstDoc->get_Document(), u"src_place"));
    dstDoc->Write(u" after");
    
    ASSERT_EQ(u"Before  after", dstDoc->get_Document()->GetText().TrimEnd(System::MakeObject<System::Array<char16_t>>(0)));
    
    // Insert source document into destination inline.
    dstDoc->MoveToBookmark(u"src_place");
    dstDoc->InsertDocumentInline(srcDoc->get_Document(), Aspose::Words::ImportFormatMode::UseDestinationStyles, System::MakeObject<Aspose::Words::ImportFormatOptions>());
    
    ASSERT_EQ(u"Before [src content] after", dstDoc->get_Document()->GetText().TrimEnd(System::MakeObject<System::Array<char16_t>>(0)));
    //ExEnd:InsertDocumentInline
}

namespace gtest_test
{

TEST_F(ExDocument, InsertDocumentInline)
{
    s_instance->InsertDocumentInline();
}

} // namespace gtest_test

void ExDocument::SaveDocumentToStream(Aspose::Words::SaveFormat saveFormat)
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Lorem ipsum");
    
    {
        System::SharedPtr<System::IO::Stream> stream = System::MakeObject<System::IO::MemoryStream>();
        if (saveFormat == Aspose::Words::SaveFormat::HtmlFixed)
        {
            auto saveOptions = System::MakeObject<Aspose::Words::Saving::HtmlFixedSaveOptions>();
            saveOptions->set_ExportEmbeddedCss(true);
            saveOptions->set_ExportEmbeddedFonts(true);
            saveOptions->set_SaveFormat(saveFormat);
            
            doc->Save(stream, saveOptions);
        }
        else if (saveFormat == Aspose::Words::SaveFormat::XamlFixed)
        {
            auto saveOptions = System::MakeObject<Aspose::Words::Saving::XamlFixedSaveOptions>();
            saveOptions->set_ResourcesFolder(get_ArtifactsDir());
            saveOptions->set_SaveFormat(saveFormat);
            
            doc->Save(stream, saveOptions);
        }
        else
        {
            doc->Save(stream, saveFormat);
        }
    }
}

namespace gtest_test
{

using ExDocument_SaveDocumentToStream_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExDocument::SaveDocumentToStream)>::type;

struct ExDocument_SaveDocumentToStream : public ExDocument, public Aspose::Words::ApiExamples::ExDocument, public ::testing::WithParamInterface<ExDocument_SaveDocumentToStream_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::SaveFormat::Doc),
            std::make_tuple(Aspose::Words::SaveFormat::Dot),
            std::make_tuple(Aspose::Words::SaveFormat::Docx),
            std::make_tuple(Aspose::Words::SaveFormat::Docm),
            std::make_tuple(Aspose::Words::SaveFormat::Dotx),
            std::make_tuple(Aspose::Words::SaveFormat::Dotm),
            std::make_tuple(Aspose::Words::SaveFormat::FlatOpc),
            std::make_tuple(Aspose::Words::SaveFormat::FlatOpcMacroEnabled),
            std::make_tuple(Aspose::Words::SaveFormat::FlatOpcTemplate),
            std::make_tuple(Aspose::Words::SaveFormat::FlatOpcTemplateMacroEnabled),
            std::make_tuple(Aspose::Words::SaveFormat::Rtf),
            std::make_tuple(Aspose::Words::SaveFormat::WordML),
            std::make_tuple(Aspose::Words::SaveFormat::Pdf),
            std::make_tuple(Aspose::Words::SaveFormat::Xps),
            std::make_tuple(Aspose::Words::SaveFormat::XamlFixed),
            std::make_tuple(Aspose::Words::SaveFormat::Svg),
            std::make_tuple(Aspose::Words::SaveFormat::HtmlFixed),
            std::make_tuple(Aspose::Words::SaveFormat::OpenXps),
            std::make_tuple(Aspose::Words::SaveFormat::Ps),
            std::make_tuple(Aspose::Words::SaveFormat::Pcl),
            std::make_tuple(Aspose::Words::SaveFormat::Html),
            std::make_tuple(Aspose::Words::SaveFormat::Mhtml),
            std::make_tuple(Aspose::Words::SaveFormat::Epub),
            std::make_tuple(Aspose::Words::SaveFormat::Azw3),
            std::make_tuple(Aspose::Words::SaveFormat::Mobi),
            std::make_tuple(Aspose::Words::SaveFormat::Odt),
            std::make_tuple(Aspose::Words::SaveFormat::Ott),
            std::make_tuple(Aspose::Words::SaveFormat::Text),
            std::make_tuple(Aspose::Words::SaveFormat::XamlFlow),
            std::make_tuple(Aspose::Words::SaveFormat::XamlFlowPack),
            std::make_tuple(Aspose::Words::SaveFormat::Markdown),
            std::make_tuple(Aspose::Words::SaveFormat::Xlsx),
            std::make_tuple(Aspose::Words::SaveFormat::Tiff),
            std::make_tuple(Aspose::Words::SaveFormat::Png),
            std::make_tuple(Aspose::Words::SaveFormat::Bmp),
            std::make_tuple(Aspose::Words::SaveFormat::Emf),
            std::make_tuple(Aspose::Words::SaveFormat::Jpeg),
            std::make_tuple(Aspose::Words::SaveFormat::Gif),
            std::make_tuple(Aspose::Words::SaveFormat::Eps),
        };
    }
};

TEST_P(ExDocument_SaveDocumentToStream, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SaveDocumentToStream(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_SaveDocumentToStream, ::testing::ValuesIn(ExDocument_SaveDocumentToStream::TestCases()));

} // namespace gtest_test

void ExDocument::HasMacros()
{
    //ExStart:HasMacros
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:FileFormatInfo.HasMacros
    //ExSummary:Shows how to check VBA macro presence without loading document.
    System::SharedPtr<Aspose::Words::FileFormatInfo> fileFormatInfo = Aspose::Words::FileFormatUtil::DetectFileFormat(get_MyDir() + u"Macro.docm");
    ASSERT_TRUE(fileFormatInfo->get_HasMacros());
    //ExEnd:HasMacros
}

namespace gtest_test
{

TEST_F(ExDocument, HasMacros)
{
    s_instance->HasMacros();
}

} // namespace gtest_test

void ExDocument::PunctuationKerning()
{
    //ExStart
    //ExFor:Document.PunctuationKerning
    //ExSummary:Shows how to work with kerning applies to both Latin text and punctuation.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    ASSERT_TRUE(doc->get_PunctuationKerning());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, PunctuationKerning)
{
    s_instance->PunctuationKerning();
}

} // namespace gtest_test

void ExDocument::RemoveBlankPages()
{
    //ExStart
    //ExFor:Document.RemoveBlankPages
    //ExSummary:Shows how to remove blank pages from the document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Blank pages.docx");
    ASSERT_EQ(2, doc->get_PageCount());
    doc->RemoveBlankPages();
    doc->UpdatePageLayout();
    ASSERT_EQ(1, doc->get_PageCount());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, RemoveBlankPages)
{
    s_instance->RemoveBlankPages();
}

} // namespace gtest_test

void ExDocument::ExtractPagesWithOptions()
{
    //ExStart:ExtractPagesWithOptions
    //GistId:571cc6e23284a2ec075d15d4c32e3bbf
    //ExFor:Document.ExtractPages(int, int)
    //ExFor:PageExtractOptions
    //ExFor:PageExtractOptions.UpdatePageStartingNumber
    //ExFor:PageExtractOptions.UnlinkPagesNumberFields
    //ExSummary:Show how to reset the initial page numbering and save the NUMPAGE field.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Page fields.docx");
    
    // Default behavior:
    // The extracted page numbering is the same as in the original document, as if we had selected "Print 2 pages" in MS Word.
    // The start page will be set to 2 and the field indicating the number of pages will be removed
    // and replaced with a constant value equal to the number of pages.
    System::SharedPtr<Aspose::Words::Document> extractedDoc1 = doc->ExtractPages(1, 1);
    extractedDoc1->Save(get_ArtifactsDir() + u"Document.ExtractPagesWithOptions.Default.docx");
    
    ASSERT_EQ(1, extractedDoc1->get_Range()->get_Fields()->get_Count());
    //ExSkip
    
    // Altered behavior:
    // The extracted page numbering is reset and a new one begins,
    // as if we had copied the contents of the second page and pasted it into a new document.
    // The start page will be set to 1 and the field indicating the number of pages will be left unchanged
    // and will show the current number of pages.
    auto extractOptions = System::MakeObject<Aspose::Words::PageExtractOptions>();
    extractOptions->set_UpdatePageStartingNumber(false);
    extractOptions->set_UnlinkPagesNumberFields(false);
    System::SharedPtr<Aspose::Words::Document> extractedDoc2 = doc->ExtractPages(1, 1, extractOptions);
    extractedDoc2->Save(get_ArtifactsDir() + u"Document.ExtractPagesWithOptions.Options.docx");
    //ExEnd:ExtractPagesWithOptions
    
    ASSERT_EQ(2, extractedDoc2->get_Range()->get_Fields()->get_Count());
}

namespace gtest_test
{

TEST_F(ExDocument, ExtractPagesWithOptions)
{
    s_instance->ExtractPagesWithOptions();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
