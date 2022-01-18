#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/BorderCollection.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Comment.h>
#include <Aspose.Words.Cpp/Comparing/CompareOptions.h>
#include <Aspose.Words.Cpp/Comparing/ComparisonTargetType.h>
#include <Aspose.Words.Cpp/Comparing/Granularity.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/ControlChar.h>
#include <Aspose.Words.Cpp/ConvertUtil.h>
#include <Aspose.Words.Cpp/DigitalSignatures/CertificateHolder.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignature.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignatureType.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Fields/FieldDate.h>
#include <Aspose.Words.Cpp/Fields/FieldEnd.h>
#include <Aspose.Words.Cpp/Fields/FieldSeparator.h>
#include <Aspose.Words.Cpp/Fields/FieldStart.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Fields/FormField.h>
#include <Aspose.Words.Cpp/Fields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/FileFormatInfo.h>
#include <Aspose.Words.Cpp/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Fonts/FontInfoCollection.h>
#include <Aspose.Words.Cpp/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Fonts/FontSourceBase.h>
#include <Aspose.Words.Cpp/Fonts/MemoryFontSource.h>
#include <Aspose.Words.Cpp/Framesets/Frameset.h>
#include <Aspose.Words.Cpp/Framesets/FramesetCollection.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/INodeChangingCallback.h>
#include <Aspose.Words.Cpp/ImageWatermarkOptions.h>
#include <Aspose.Words.Cpp/ImportFormatMode.h>
#include <Aspose.Words.Cpp/ImportFormatOptions.h>
#include <Aspose.Words.Cpp/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Layout/CommentDisplayMode.h>
#include <Aspose.Words.Cpp/Layout/LayoutOptions.h>
#include <Aspose.Words.Cpp/Layout/RevisionColor.h>
#include <Aspose.Words.Cpp/Layout/RevisionOptions.h>
#include <Aspose.Words.Cpp/LineStyle.h>
#include <Aspose.Words.Cpp/Lists/List.h>
#include <Aspose.Words.Cpp/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Lists/ListLabel.h>
#include <Aspose.Words.Cpp/LoadFormat.h>
#include <Aspose.Words.Cpp/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Markup/CustomPart.h>
#include <Aspose.Words.Cpp/Markup/CustomPartCollection.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeChangingArgs.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Notes/Footnote.h>
#include <Aspose.Words.Cpp/Notes/FootnoteOptions.h>
#include <Aspose.Words.Cpp/Notes/FootnoteType.h>
#include <Aspose.Words.Cpp/Orientation.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/ProtectionType.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Rendering/ThumbnailGeneratingOptions.h>
#include <Aspose.Words.Cpp/Replacing/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Revision.h>
#include <Aspose.Words.Cpp/RevisionCollection.h>
#include <Aspose.Words.Cpp/RevisionGroup.h>
#include <Aspose.Words.Cpp/RevisionGroupCollection.h>
#include <Aspose.Words.Cpp/RevisionType.h>
#include <Aspose.Words.Cpp/RevisionsView.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Saving/SaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Saving/TxtSaveOptions.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/Settings/HyphenationOptions.h>
#include <Aspose.Words.Cpp/Settings/WriteProtection.h>
#include <Aspose.Words.Cpp/Shading.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/StyleType.h>
#include <Aspose.Words.Cpp/SubDocument.h>
#include <Aspose.Words.Cpp/TableStyle.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/TextWatermarkOptions.h>
#include <Aspose.Words.Cpp/Vba/VbaModule.h>
#include <Aspose.Words.Cpp/Vba/VbaModuleCollection.h>
#include <Aspose.Words.Cpp/Vba/VbaProject.h>
#include <Aspose.Words.Cpp/Watermark.h>
#include <Aspose.Words.Cpp/WatermarkLayout.h>
#include <Aspose.Words.Cpp/WatermarkType.h>
#include <Aspose.Words.Cpp/WebExtensions/TaskPane.h>
#include <Aspose.Words.Cpp/WebExtensions/TaskPaneCollection.h>
#include <Aspose.Words.Cpp/WebExtensions/TaskPaneDockState.h>
#include <Aspose.Words.Cpp/WebExtensions/WebExtension.h>
#include <Aspose.Words.Cpp/WebExtensions/WebExtensionBinding.h>
#include <Aspose.Words.Cpp/WebExtensions/WebExtensionBindingCollection.h>
#include <Aspose.Words.Cpp/WebExtensions/WebExtensionBindingType.h>
#include <Aspose.Words.Cpp/WebExtensions/WebExtensionProperty.h>
#include <Aspose.Words.Cpp/WebExtensions/WebExtensionPropertyCollection.h>
#include <Aspose.Words.Cpp/WebExtensions/WebExtensionReference.h>
#include <Aspose.Words.Cpp/WebExtensions/WebExtensionStoreType.h>
#include <drawing/color.h>
#include <drawing/image.h>
#include <drawing/imaging/image_format.h>
#include <drawing/size.h>
#include <net/http_status_code.h>
#include <net/web_client.h>
#include <security/cryptography/x509_certificates/x500_distinguished_name.h>
#include <security/cryptography/x509_certificates/x509_certificate_2.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ienumerator.h>
#include <system/collections/list.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/globalization/culture_info.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/file_mode.h>
#include <system/io/file_not_found_exception.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/io/stream.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/regularexpressions/match_collection.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/string_builder.h>
#include <system/timespan.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Comparing;
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
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Vba;
using namespace Aspose::Words::WebExtensions;

namespace ApiExamples {

class ExDocument : public ApiExampleBase
{
public:
    void Constructor()
    {
        //ExStart
        //ExFor:Document.#ctor()
        //ExFor:Document.#ctor(String,LoadOptions)
        //ExSummary:Shows how to create and load documents.
        // There are two ways of creating a Document object using Aspose.Words.
        // 1 -  Create a blank document:
        auto doc = MakeObject<Document>();

        // New Document objects by default come with the minimal set of nodes
        // required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Hello world!"));

        // 2 -  Load a document that exists in the local file system:
        doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Loaded documents will have contents that we can access and edit.
        ASSERT_EQ(u"Hello World!", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText().Trim());

        // Some operations that need to occur during loading, such as using a password to decrypt a document,
        // can be done by passing a LoadOptions object when loading the document.
        doc = MakeObject<Document>(MyDir + u"Encrypted.docx", MakeObject<LoadOptions>(u"docPassword"));

        ASSERT_EQ(u"Test encrypted document.", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText().Trim());
        //ExEnd
    }

    void LoadFromStream()
    {
        //ExStart
        //ExFor:Document.#ctor(Stream)
        //ExSummary:Shows how to load a document using a stream.
        {
            SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(MyDir + u"Document.docx");
            auto doc = MakeObject<Document>(stream);

            ASSERT_EQ(u"Hello World!\r\rHello Word!\r\r\rHello World!", doc->GetText().Trim());
        }
        //ExEnd
    }

    void LoadFromWeb()
    {
        //ExStart
        //ExFor:Document.#ctor(Stream)
        //ExSummary:Shows how to load a document from a URL.
        // Create a URL that points to a Microsoft Word document.
        const String url = u"https://omextemplates.content.office.net/support/templates/en-us/tf16402488.dotx";

        // Download the document into a byte array, then load that array into a document using a memory stream.
        {
            auto webClient = MakeObject<System::Net::WebClient>();
            ArrayPtr<uint8_t> dataBytes = webClient->DownloadData(url);

            {
                auto byteStream = MakeObject<System::IO::MemoryStream>(dataBytes);
                auto doc = MakeObject<Document>(byteStream);

                // At this stage, we can read and edit the document's contents and then save it to the local file system.
                ASSERT_EQ(String(u"Use this section to highlight your relevant passions, activities, and how you like to give back. ") +
                              u"It’s good to include Leadership and volunteer experiences here. " +
                              u"Or show off important extras like publications, certifications, languages and more.",
                          doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(4)->GetText().Trim());

                doc->Save(ArtifactsDir + u"Document.LoadFromWeb.docx");
            }
        }
        //ExEnd

        TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, url);
    }

    void ConvertToPdf()
    {
        //ExStart
        //ExFor:Document.#ctor(String)
        //ExFor:Document.Save(String)
        //ExSummary:Shows how to open a document and convert it to .PDF.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        doc->Save(ArtifactsDir + u"Document.ConvertToPdf.pdf");
        //ExEnd
    }

    void SaveToImageStream()
    {
        //ExStart
        //ExFor:Document.Save(Stream, SaveFormat)
        //ExSummary:Shows how to save a document to an image via stream, and then read the image from that stream.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Times New Roman");
        builder->get_Font()->set_Size(24);
        builder->Writeln(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        builder->InsertImage(ImageDir + u"Logo.jpg");

        {
            auto stream = MakeObject<System::IO::MemoryStream>();
            doc->Save(stream, SaveFormat::Bmp);

            stream->set_Position(0);

            // Read the stream back into an image.
            {
                SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromStream(stream);
                ASPOSE_ASSERT_EQ(System::Drawing::Imaging::ImageFormat::get_Bmp(), image->get_RawFormat());
                ASSERT_EQ(816, image->get_Width());
                ASSERT_EQ(1056, image->get_Height());
            }
        }
        //ExEnd
    }

    void OpenFromStreamWithBaseUri()
    {
        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExFor:LoadOptions.#ctor()
        //ExFor:LoadOptions.BaseUri
        //ExSummary:Shows how to open an HTML document with images from a stream using a base URI.
        {
            SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(MyDir + u"Document.html");
            // Pass the URI of the base folder while loading it
            // so that any images with relative URIs in the HTML document can be found.
            auto loadOptions = MakeObject<LoadOptions>();
            loadOptions->set_BaseUri(ImageDir);

            auto doc = MakeObject<Document>(stream, loadOptions);

            // Verify that the first shape of the document contains a valid image.
            auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

            ASSERT_TRUE(shape->get_IsImage());
            ASSERT_FALSE(shape->get_ImageData()->get_ImageBytes() == nullptr);
            ASSERT_NEAR(32.0, ConvertUtil::PointToPixel(shape->get_Width()), 0.01);
            ASSERT_NEAR(32.0, ConvertUtil::PointToPixel(shape->get_Height()), 0.01);
        }
        //ExEnd
    }

    void InsertHtmlFromWebPage()
    {
        //ExStart
        //ExFor:Document.#ctor(Stream, LoadOptions)
        //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
        //ExFor:LoadFormat
        //ExSummary:Shows how save a web page as a .docx file.
        const String url = u"https://www.aspose.com/";

        {
            auto client = MakeObject<System::Net::WebClient>();
            {
                auto stream = MakeObject<System::IO::MemoryStream>(client->DownloadData(url));
                // The URL is used again as a baseUri to ensure that any relative image paths are retrieved correctly.
                auto options = MakeObject<LoadOptions>(LoadFormat::Html, u"", url);

                // Load the HTML document from stream and pass the LoadOptions object.
                auto doc = MakeObject<Document>(stream, options);

                // At this stage, we can read and edit the document's contents and then save it to the local file system.
                ASSERT_EQ(u"File Format APIs", doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_Runs()->idx_get(0)->GetText().Trim());
                //ExSkip

                doc->Save(ArtifactsDir + u"Document.InsertHtmlFromWebPage.docx");
            }
        }
        //ExEnd

        TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, url);
    }

    void LoadEncrypted()
    {
        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExFor:Document.#ctor(String,LoadOptions)
        //ExFor:LoadOptions
        //ExFor:LoadOptions.#ctor(String)
        //ExSummary:Shows how to load an encrypted Microsoft Word document.
        SharedPtr<Document> doc;

        // Aspose.Words throw an exception if we try to open an encrypted document without its password.
        ASSERT_THROW(doc = MakeObject<Document>(MyDir + u"Encrypted.docx"), IncorrectPasswordException);

        // When loading such a document, the password is passed to the document's constructor using a LoadOptions object.
        auto options = MakeObject<LoadOptions>(u"docPassword");

        // There are two ways of loading an encrypted document with a LoadOptions object.
        // 1 -  Load the document from the local file system by filename:
        doc = MakeObject<Document>(MyDir + u"Encrypted.docx", options);
        ASSERT_EQ(u"Test encrypted document.", doc->GetText().Trim());
        //ExSkip

        // 2 -  Load the document from a stream:
        {
            SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(MyDir + u"Encrypted.docx");
            doc = MakeObject<Document>(stream, options);
            ASSERT_EQ(u"Test encrypted document.", doc->GetText().Trim());
            //ExSkip
        }
        //ExEnd
    }

    void TempFolder()
    {
        //ExStart
        //ExFor:LoadOptions.TempFolder
        //ExSummary:Shows how to load a document using temporary files.
        // Note that such an approach can reduce memory usage but degrades speed
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_TempFolder(u"C:\\TempFolder\\");

        // Ensure that the directory exists and load
        System::IO::Directory::CreateDirectory_(loadOptions->get_TempFolder());

        auto doc = MakeObject<Document>(MyDir + u"Document.docx", loadOptions);
        //ExEnd
    }

    void ConvertToHtml()
    {
        //ExStart
        //ExFor:Document.Save(String,SaveFormat)
        //ExFor:SaveFormat
        //ExSummary:Shows how to convert from DOCX to HTML format.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        doc->Save(ArtifactsDir + u"Document.ConvertToHtml.html", SaveFormat::Html);
        //ExEnd
    }

    void ConvertToMhtml()
    {
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        doc->Save(ArtifactsDir + u"Document.ConvertToMhtml.mht");
    }

    void ConvertToTxt()
    {
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        doc->Save(ArtifactsDir + u"Document.ConvertToTxt.txt");
    }

    void ConvertToEpub()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        doc->Save(ArtifactsDir + u"Document.ConvertToEpub.epub");
    }

    void SaveToStream()
    {
        //ExStart
        //ExFor:Document.Save(Stream,SaveFormat)
        //ExSummary:Shows how to save a document to a stream.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        {
            auto dstStream = MakeObject<System::IO::MemoryStream>();
            doc->Save(dstStream, SaveFormat::Docx);

            // Verify that the stream contains the document.
            ASSERT_EQ(u"Hello World!\r\rHello Word!\r\r\rHello World!", MakeObject<Document>(dstStream)->GetText().Trim());
        }
        //ExEnd
    }

    //ExStart
    //ExFor:INodeChangingCallback
    //ExFor:INodeChangingCallback.NodeInserting
    //ExFor:INodeChangingCallback.NodeInserted
    //ExFor:INodeChangingCallback.NodeRemoving
    //ExFor:INodeChangingCallback.NodeRemoved
    //ExFor:NodeChangingArgs
    //ExFor:NodeChangingArgs.Node
    //ExFor:DocumentBase.NodeChangingCallback
    //ExSummary:Shows how customize node changing with a callback.
    void FontChangeViaCallback()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set the node changing callback to custom implementation,
        // then add/remove nodes to get it to generate a log.
        auto callback = MakeObject<ExDocument::HandleNodeChangingFontChanger>();
        doc->set_NodeChangingCallback(callback);

        builder->Writeln(u"Hello world!");
        builder->Writeln(u"Hello again!");
        builder->InsertField(u" HYPERLINK \"https://www.google.com/\" ");
        builder->InsertShape(ShapeType::Rectangle, 300, 300);

        doc->get_Range()->get_Fields()->idx_get(0)->Remove();

        std::cout << callback->GetLog() << std::endl;
        TestFontChangeViaCallback(callback->GetLog());
        //ExSkip
    }

    /// <summary>
    /// Logs the date and time of each node insertion and removal.
    /// Sets a custom font name/size for the text contents of Run nodes.
    /// </summary>
    class HandleNodeChangingFontChanger : public INodeChangingCallback
    {
    public:
        String GetLog()
        {
            return mLog->ToString();
        }

        HandleNodeChangingFontChanger() : mLog(MakeObject<System::Text::StringBuilder>())
        {
        }

    private:
        SharedPtr<System::Text::StringBuilder> mLog;

        void NodeInserted(SharedPtr<NodeChangingArgs> args) override
        {
            mLog->AppendLine(String::Format(u"\tType:\t{0}", args->get_Node()->get_NodeType()));
            mLog->AppendLine(String::Format(u"\tHash:\t{0}", System::ObjectExt::GetHashCode(args->get_Node())));

            if (args->get_Node()->get_NodeType() == NodeType::Run)
            {
                SharedPtr<Aspose::Words::Font> font = (System::DynamicCast<Run>(args->get_Node()))->get_Font();
                mLog->Append(String::Format(u"\tFont:\tChanged from \"{0}\" {1}pt", font->get_Name(), font->get_Size()));

                font->set_Size(24);
                font->set_Name(u"Arial");

                mLog->AppendLine(String::Format(u" to \"{0}\" {1}pt", font->get_Name(), font->get_Size()));
                mLog->AppendLine(String::Format(u"\tContents:\n\t\t\"{0}\"", args->get_Node()->GetText()));
            }
        }

        void NodeInserting(SharedPtr<NodeChangingArgs> args) override
        {
            mLog->AppendLine(String::Format(u"\n{0:dd/MM/yyyy HH:mm:ss:fff}\tNode insertion:", System::DateTime::get_Now()));
        }

        void NodeRemoved(SharedPtr<NodeChangingArgs> args) override
        {
            mLog->AppendLine(String::Format(u"\tType:\t{0}", args->get_Node()->get_NodeType()));
            mLog->AppendLine(String::Format(u"\tHash code:\t{0}", System::ObjectExt::GetHashCode(args->get_Node())));
        }

        void NodeRemoving(SharedPtr<NodeChangingArgs> args) override
        {
            mLog->AppendLine(String::Format(u"\n{0:dd/MM/yyyy HH:mm:ss:fff}\tNode removal:", System::DateTime::get_Now()));
        }
    };
    //ExEnd

    static void TestFontChangeViaCallback(String log)
    {
        ASSERT_EQ(10, System::Text::RegularExpressions::Regex::Matches(log, u"insertion")->get_Count());
        ASSERT_EQ(5, System::Text::RegularExpressions::Regex::Matches(log, u"removal")->get_Count());
    }

    void AppendDocument()
    {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode)
        //ExSummary:Shows how to append a document to the end of another document.
        auto srcDoc = MakeObject<Document>();
        srcDoc->get_FirstSection()->get_Body()->AppendParagraph(u"Source document text. ");

        auto dstDoc = MakeObject<Document>();
        dstDoc->get_FirstSection()->get_Body()->AppendParagraph(u"Destination document text. ");

        // Append the source document to the destination document while preserving its formatting,
        // then save the source document to the local file system.
        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);
        ASSERT_EQ(2, dstDoc->get_Sections()->get_Count());
        //ExSkip

        dstDoc->Save(ArtifactsDir + u"Document.AppendDocument.docx");
        //ExEnd

        String outDocText = MakeObject<Document>(ArtifactsDir + u"Document.AppendDocument.docx")->GetText();

        ASSERT_TRUE(outDocText.StartsWith(dstDoc->GetText()));
        ASSERT_TRUE(outDocText.EndsWith(srcDoc->GetText()));
    }

    void AppendDocumentFromAutomation()
    {
        auto doc = MakeObject<Document>();

        // We should call this method to clear this document of any existing content.
        doc->RemoveAllChildren();

        const int recordCount = 5;
        for (int i = 1; i <= recordCount; i++)
        {
            auto srcDoc = MakeObject<Document>();

            ASSERT_THROW(static_cast<std::function<bool()>>([&srcDoc]() -> bool { return srcDoc == MakeObject<Document>(u"C:\\DetailsList.doc"); })(),
                         System::IO::FileNotFoundException);

            // Append the source document at the end of the destination document.
            doc->AppendDocument(srcDoc, ImportFormatMode::UseDestinationStyles);

            // Automation required you to insert a new section break at this point, however, in Aspose.Words we
            // do not need to do anything here as the appended document is imported as separate sections already

            // Unlink all headers/footers in this section from the previous section headers/footers
            // if this is the second document or above being appended.
            if (i > 1)
            {
                ASSERT_THROW(doc->get_Sections()->idx_get(i)->get_HeadersFooters()->LinkToPrevious(false), System::NullReferenceException);
            }
        }
    }

    void ImportList(bool isKeepSourceNumbering)
    {
        //ExStart
        //ExFor:ImportFormatOptions.KeepSourceNumbering
        //ExSummary:Shows how to import a document with numbered lists.
        auto srcDoc = MakeObject<Document>(MyDir + u"List source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"List destination.docx");

        ASSERT_EQ(4, dstDoc->get_Lists()->get_Count());

        auto options = MakeObject<ImportFormatOptions>();

        // If there is a clash of list styles, apply the list format of the source document.
        // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
        // Set the "KeepSourceNumbering" property to "true" import all clashing
        // list style numbering with the same appearance that it had in the source document.
        options->set_KeepSourceNumbering(isKeepSourceNumbering);

        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting, options);
        dstDoc->UpdateListLabels();

        ASSERT_EQ(isKeepSourceNumbering ? 5 : 4, dstDoc->get_Lists()->get_Count());
        //ExEnd
    }

    void KeepSourceNumberingSameListIds()
    {
        //ExStart
        //ExFor:ImportFormatOptions.KeepSourceNumbering
        //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how resolve a clash when importing documents that have lists with the same list definition identifier.
        auto srcDoc = MakeObject<Document>(MyDir + u"List with the same definition identifier - source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"List with the same definition identifier - destination.docx");

        // Set the "KeepSourceNumbering" property to "true" to apply a different list definition ID
        // to identical styles as Aspose.Words imports them into destination documents.
        auto importFormatOptions = MakeObject<ImportFormatOptions>();
        importFormatOptions->set_KeepSourceNumbering(true);

        dstDoc->AppendDocument(srcDoc, ImportFormatMode::UseDestinationStyles, importFormatOptions);
        dstDoc->UpdateListLabels();
        //ExEnd

        String paraText = dstDoc->get_Sections()->idx_get(1)->get_Body()->get_LastParagraph()->GetText();

        ASSERT_TRUE(paraText.StartsWith(u"13->13")) << (paraText);
        ASSERT_EQ(u"1.", dstDoc->get_Sections()->idx_get(1)->get_Body()->get_LastParagraph()->get_ListLabel()->get_LabelString());
    }

    void MergePastedLists()
    {
        //ExStart
        //ExFor:ImportFormatOptions.MergePastedLists
        //ExSummary:Shows how to merge lists from a documents.
        auto srcDoc = MakeObject<Document>(MyDir + u"List item.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"List destination.docx");

        auto options = MakeObject<ImportFormatOptions>();
        options->set_MergePastedLists(true);

        // Set the "MergePastedLists" property to "true" pasted lists will be merged with surrounding lists.
        dstDoc->AppendDocument(srcDoc, ImportFormatMode::UseDestinationStyles, options);

        dstDoc->Save(ArtifactsDir + u"Document.MergePastedLists.docx");
        //ExEnd
    }

    void ValidateIndividualDocumentSignatures()
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
        auto doc = MakeObject<Document>(MyDir + u"Digitally signed.docx");

        for (const auto& signature : doc->get_DigitalSignatures())
        {
            std::cout << (signature->get_IsValid() ? String(u"Valid") : String(u"Invalid")) << " signature: " << std::endl;
            std::cout << "\tReason:\t" << signature->get_Comments() << std::endl;
            std::cout << String::Format(u"\tType:\t{0}", signature->get_SignatureType()) << std::endl;
            std::cout << "\tSign time:\t" << signature->get_SignTime() << std::endl;
            std::cout << "\tSubject name:\t" << signature->get_CertificateHolder()->get_Certificate()->get_SubjectName() << std::endl;
            std::cout << "\tIssuer name:\t" << signature->get_CertificateHolder()->get_Certificate()->get_IssuerName()->get_Name() << std::endl;
            std::cout << std::endl;
        }
        //ExEnd

        ASSERT_EQ(1, doc->get_DigitalSignatures()->get_Count());

        SharedPtr<DigitalSignature> digitalSig = doc->get_DigitalSignatures()->idx_get(0);

        ASSERT_TRUE(digitalSig->get_IsValid());
        ASSERT_EQ(u"Test Sign", digitalSig->get_Comments());
        ASSERT_EQ(u"XmlDsig", System::ObjectExt::ToString(digitalSig->get_SignatureType()));
        ASSERT_TRUE(digitalSig->get_CertificateHolder()->get_Certificate()->get_Subject().Contains(u"Aspose Pty Ltd"));
        ASSERT_TRUE(digitalSig->get_CertificateHolder()->get_Certificate()->get_IssuerName()->get_Name() != nullptr &&
                    digitalSig->get_CertificateHolder()->get_Certificate()->get_IssuerName()->get_Name().Contains(u"VeriSign"));
    }

    void DigitalSignature_()
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
        ASSERT_FALSE(FileFormatUtil::DetectFileFormat(MyDir + u"Document.docx")->get_HasDigitalSignature());

        // Create a CertificateHolder object from a PKCS12 file, which we will use to sign the document.
        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw", nullptr);

        // There are two ways of saving a signed copy of a document to the local file system:
        // 1 - Designate a document by a local system filename and save a signed copy at a location specified by another filename.
        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_SignTime(System::DateTime::get_Now());

        DigitalSignatureUtil::Sign(MyDir + u"Document.docx", ArtifactsDir + u"Document.DigitalSignature.docx", certificateHolder, signOptions);

        ASSERT_TRUE(FileFormatUtil::DetectFileFormat(ArtifactsDir + u"Document.DigitalSignature.docx")->get_HasDigitalSignature());

        // 2 - Take a document from a stream and save a signed copy to another stream.
        {
            auto inDoc = MakeObject<System::IO::FileStream>(MyDir + u"Document.docx", System::IO::FileMode::Open);
            {
                auto outDoc = MakeObject<System::IO::FileStream>(ArtifactsDir + u"Document.DigitalSignature.docx", System::IO::FileMode::Create);
                DigitalSignatureUtil::Sign(inDoc, outDoc, certificateHolder);
            }
        }

        ASSERT_TRUE(FileFormatUtil::DetectFileFormat(ArtifactsDir + u"Document.DigitalSignature.docx")->get_HasDigitalSignature());

        // Please verify that all of the document's digital signatures are valid and check their details.
        auto signedDoc = MakeObject<Document>(ArtifactsDir + u"Document.DigitalSignature.docx");
        SharedPtr<DigitalSignatureCollection> digitalSignatureCollection = signedDoc->get_DigitalSignatures();

        ASSERT_TRUE(digitalSignatureCollection->get_IsValid());
        ASSERT_EQ(1, digitalSignatureCollection->get_Count());
        ASSERT_EQ(DigitalSignatureType::XmlDsig, digitalSignatureCollection->idx_get(0)->get_SignatureType());
        ASSERT_EQ(u"CN=Morzal.Me", signedDoc->get_DigitalSignatures()->idx_get(0)->get_IssuerName());
        ASSERT_EQ(u"CN=Morzal.Me", signedDoc->get_DigitalSignatures()->idx_get(0)->get_SubjectName());
        //ExEnd
    }

    void AppendAllDocumentsInFolder()
    {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode)
        //ExSummary:Shows how to append all the documents in a folder to the end of a template document.
        auto dstDoc = MakeObject<Document>();

        auto builder = MakeObject<DocumentBuilder>(dstDoc);
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);
        builder->Writeln(u"Template Document");
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Normal);
        builder->Writeln(u"Some content here");
        ASSERT_EQ(5, dstDoc->get_Styles()->get_Count());
        //ExSkip
        ASSERT_EQ(1, dstDoc->get_Sections()->get_Count());
        //ExSkip

        // Append all unencrypted documents with the .doc extension
        // from our local file system directory to the base document.
        auto isDocFile = [](String item)
        {
            return item.EndsWith(u".doc");
        };

        SharedPtr<System::Collections::Generic::List<String>> docFiles = System::IO::Directory::GetFiles(MyDir, u"*.doc")->LINQ_Where(isDocFile)->LINQ_ToList();
        for (const auto& fileName : docFiles)
        {
            SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(fileName);
            if (info->get_IsEncrypted())
            {
                continue;
            }

            auto srcDoc = MakeObject<Document>(fileName);
            dstDoc->AppendDocument(srcDoc, ImportFormatMode::UseDestinationStyles);
        }

        dstDoc->Save(ArtifactsDir + u"Document.AppendAllDocumentsInFolder.doc");
        //ExEnd

        ASSERT_EQ(7, dstDoc->get_Styles()->get_Count());
        ASSERT_EQ(9, dstDoc->get_Sections()->get_Count());
    }

    void JoinRunsWithSameFormatting()
    {
        //ExStart
        //ExFor:Document.JoinRunsWithSameFormatting
        //ExSummary:Shows how to join runs in a document to reduce unneeded runs.
        // Open a document that contains adjacent runs of text with identical formatting,
        // which commonly occurs if we edit the same paragraph multiple times in Microsoft Word.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // If any number of these runs are adjacent with identical formatting,
        // then the document may be simplified.
        ASSERT_EQ(317, doc->GetChildNodes(NodeType::Run, true)->get_Count());

        // Combine such runs with this method and verify the number of run joins that will take place.
        ASSERT_EQ(121, doc->JoinRunsWithSameFormatting());

        // The number of joins and the number of runs we have after the join
        // should add up the number of runs we had initially.
        ASSERT_EQ(196, doc->GetChildNodes(NodeType::Run, true)->get_Count());
        //ExEnd
    }

    void DefaultTabStop()
    {
        //ExStart
        //ExFor:Document.DefaultTabStop
        //ExFor:ControlChar.Tab
        //ExFor:ControlChar.TabChar
        //ExSummary:Shows how to set a custom interval for tab stop positions.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set tab stops to appear every 72 points (1 inch).
        builder->get_Document()->set_DefaultTabStop(72);

        // Each tab character snaps the text after it to the next closest tab stop position.
        builder->Writeln(String(u"Hello") + ControlChar::Tab() + u"World!");
        builder->Writeln(String(u"Hello") + ControlChar::TabChar + u"World!");
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        ASPOSE_ASSERT_EQ(72, doc->get_DefaultTabStop());
    }

    void CloneDocument()
    {
        //ExStart
        //ExFor:Document.Clone
        //ExSummary:Shows how to deep clone a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Hello world!");

        // Cloning will produce a new document with the same contents as the original,
        // but with a unique copy of each of the original document's nodes.
        SharedPtr<Document> clone = doc->Clone();

        ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->GetText(),
                  clone->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Text());
        ASSERT_NE(System::ObjectExt::GetHashCode(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)),
                  System::ObjectExt::GetHashCode(clone->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)));
        //ExEnd
    }

    void DocumentGetTextToString()
    {
        //ExStart
        //ExFor:CompositeNode.GetText
        //ExFor:Node.ToString(SaveFormat)
        //ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
        auto doc = MakeObject<Document>();

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertField(u"MERGEFIELD Field");

        // GetText will retrieve the visible text as well as field codes and special characters.
        ASSERT_EQ(u"\u0013MERGEFIELD Field\u0014«Field»\u0015\u000c", doc->GetText());

        // ToString will give us the document's appearance if saved to a passed save format.
        ASSERT_EQ(u"«Field»\r\n", doc->ToString(SaveFormat::Text));
        //ExEnd
    }

    void DocumentByteArray()
    {
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto streamOut = MakeObject<System::IO::MemoryStream>();
        doc->Save(streamOut, SaveFormat::Docx);

        ArrayPtr<uint8_t> docBytes = streamOut->ToArray();

        auto streamIn = MakeObject<System::IO::MemoryStream>(docBytes);

        auto loadDoc = MakeObject<Document>(streamIn);
        ASSERT_EQ(doc->GetText(), loadDoc->GetText());
    }

    void ProtectUnprotect()
    {
        //ExStart
        //ExFor:Document.Protect(ProtectionType,String)
        //ExFor:Document.ProtectionType
        //ExFor:Document.Unprotect()
        //ExFor:Document.Unprotect(String)
        //ExSummary:Shows how to protect and unprotect a document.
        auto doc = MakeObject<Document>();
        doc->Protect(ProtectionType::ReadOnly, u"password");

        ASSERT_EQ(ProtectionType::ReadOnly, doc->get_ProtectionType());

        // If we open this document with Microsoft Word intending to edit it,
        // we will need to apply the password to get through the protection.
        doc->Save(ArtifactsDir + u"Document.Protect.docx");

        // Note that the protection only applies to Microsoft Word users opening our document.
        // We have not encrypted the document in any way, and we do not need the password to open and edit it programmatically.
        auto protectedDoc = MakeObject<Document>(ArtifactsDir + u"Document.Protect.docx");

        ASSERT_EQ(ProtectionType::ReadOnly, protectedDoc->get_ProtectionType());

        auto builder = MakeObject<DocumentBuilder>(protectedDoc);
        builder->Writeln(u"Text added to a protected document.");
        ASSERT_EQ(u"Text added to a protected document.", protectedDoc->get_Range()->get_Text().Trim());
        //ExSkip

        // There are two ways of removing protection from a document.
        // 1 - With no password:
        doc->Unprotect();

        ASSERT_EQ(ProtectionType::NoProtection, doc->get_ProtectionType());

        doc->Protect(ProtectionType::ReadOnly, u"NewPassword");

        ASSERT_EQ(ProtectionType::ReadOnly, doc->get_ProtectionType());

        doc->Unprotect(u"WrongPassword");

        ASSERT_EQ(ProtectionType::ReadOnly, doc->get_ProtectionType());

        // 2 - With the correct password:
        doc->Unprotect(u"NewPassword");

        ASSERT_EQ(ProtectionType::NoProtection, doc->get_ProtectionType());
        //ExEnd
    }

    void DocumentEnsureMinimum()
    {
        //ExStart
        //ExFor:Document.EnsureMinimum
        //ExSummary:Shows how to ensure that a document contains the minimal set of nodes required for editing its contents.
        // A newly created document contains one child Section, which includes one child Body and one child Paragraph.
        // We can edit the document body's contents by adding nodes such as Runs or inline Shapes to that paragraph.
        auto doc = MakeObject<Document>();
        SharedPtr<NodeCollection> nodes = doc->GetChildNodes(NodeType::Any, true);

        ASSERT_EQ(NodeType::Section, nodes->idx_get(0)->get_NodeType());
        ASPOSE_ASSERT_EQ(doc, nodes->idx_get(0)->get_ParentNode());

        ASSERT_EQ(NodeType::Body, nodes->idx_get(1)->get_NodeType());
        ASPOSE_ASSERT_EQ(nodes->idx_get(0), nodes->idx_get(1)->get_ParentNode());

        ASSERT_EQ(NodeType::Paragraph, nodes->idx_get(2)->get_NodeType());
        ASPOSE_ASSERT_EQ(nodes->idx_get(1), nodes->idx_get(2)->get_ParentNode());

        // This is the minimal set of nodes that we need to be able to edit the document.
        // We will no longer be able to edit the document if we remove any of them.
        doc->RemoveAllChildren();

        ASSERT_EQ(0, doc->GetChildNodes(NodeType::Any, true)->get_Count());

        // Call this method to make sure that the document has at least those three nodes so we can edit it again.
        doc->EnsureMinimum();

        ASSERT_EQ(NodeType::Section, nodes->idx_get(0)->get_NodeType());
        ASSERT_EQ(NodeType::Body, nodes->idx_get(1)->get_NodeType());
        ASSERT_EQ(NodeType::Paragraph, nodes->idx_get(2)->get_NodeType());

        (System::DynamicCast<Paragraph>(nodes->idx_get(2)))->get_Runs()->Add(MakeObject<Run>(doc, u"Hello world!"));
        //ExEnd

        ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
    }

    void RemoveMacrosFromDocument()
    {
        //ExStart
        //ExFor:Document.RemoveMacros
        //ExSummary:Shows how to remove all macros from a document.
        auto doc = MakeObject<Document>(MyDir + u"Macro.docm");

        ASSERT_TRUE(doc->get_HasMacros());
        ASSERT_EQ(u"Project", doc->get_VbaProject()->get_Name());

        // Remove the document's VBA project, along with all its macros.
        doc->RemoveMacros();

        ASSERT_FALSE(doc->get_HasMacros());
        ASSERT_TRUE(doc->get_VbaProject() == nullptr);
        //ExEnd
    }

    void GetPageCount()
    {
        //ExStart
        //ExFor:Document.PageCount
        //ExSummary:Shows how to count the number of pages in the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Page 1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Write(u"Page 2");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Write(u"Page 3");

        // Verify the expected page count of the document.
        ASSERT_EQ(3, doc->get_PageCount());

        // Getting the PageCount property invoked the document's page layout to calculate the value.
        // This operation will not need to be re-done when rendering the document to a fixed page save format,
        // such as .pdf. So you can save some time, especially with more complex documents.
        doc->Save(ArtifactsDir + u"Document.GetPageCount.pdf");
        //ExEnd
    }

    void GetUpdatedPageProperties()
    {
        //ExStart
        //ExFor:Document.UpdateWordCount()
        //ExFor:Document.UpdateWordCount(Boolean)
        //ExFor:BuiltInDocumentProperties.Characters
        //ExFor:BuiltInDocumentProperties.Words
        //ExFor:BuiltInDocumentProperties.Paragraphs
        //ExFor:BuiltInDocumentProperties.Lines
        //ExSummary:Shows how to update all list labels in a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") +
                         u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        builder->Write(String(u"Ut enim ad minim veniam, ") + u"quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

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

    void TableStyleToDirectFormatting()
    {
        //ExStart
        //ExFor:CompositeNode.GetChild
        //ExFor:Document.ExpandTableStylesToDirectFormatting
        //ExSummary:Shows how to apply the properties of a table's style directly to the table's elements.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Hello world!");
        builder->EndTable();

        auto tableStyle = System::DynamicCast<TableStyle>(doc->get_Styles()->Add(StyleType::Table, u"MyTableStyle1"));
        tableStyle->set_RowStripe(3);
        tableStyle->set_CellSpacing(5);
        tableStyle->get_Shading()->set_BackgroundPatternColor(System::Drawing::Color::get_AntiqueWhite());
        tableStyle->get_Borders()->set_Color(System::Drawing::Color::get_Blue());
        tableStyle->get_Borders()->set_LineStyle(LineStyle::DotDash);

        table->set_Style(tableStyle);

        // This method concerns table style properties such as the ones we set above.
        doc->ExpandTableStylesToDirectFormatting();

        doc->Save(ArtifactsDir + u"Document.TableStyleToDirectFormatting.docx");
        //ExEnd
    }

    void UpdateTableLayout()
    {
        //ExStart
        //ExFor:Document.UpdateTableLayout
        //ExSummary:Shows how to preserve a table's layout when saving to .txt.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Cell 1");
        builder->InsertCell();
        builder->Write(u"Cell 2");
        builder->InsertCell();
        builder->Write(u"Cell 3");
        builder->EndTable();

        // Use a TxtSaveOptions object to preserve the table's layout when converting the document to plaintext.
        auto options = MakeObject<TxtSaveOptions>();
        options->set_PreserveTableLayout(true);

        // Previewing the appearance of the document in .txt form shows that the table will not be represented accurately.
        ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_Width());
        ASSERT_EQ(u"CCC\r\neee\r\nlll\r\nlll\r\n   \r\n123\r\n\r\n", doc->ToString(options));

        // We can call UpdateTableLayout() to fix some of these issues.
        doc->UpdateTableLayout();

        ASSERT_EQ(u"Cell 1                                       Cell 2                                       Cell 3\r\n\r\n", doc->ToString(options));
        ASSERT_NEAR(155.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_Width(), 2.f);
        //ExEnd
    }

    void GetOriginalFileInfo()
    {
        //ExStart
        //ExFor:Document.OriginalFileName
        //ExFor:Document.OriginalLoadFormat
        //ExSummary:Shows how to retrieve details of a document's load operation.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        ASSERT_EQ(MyDir + u"Document.docx", doc->get_OriginalFileName());
        ASSERT_EQ(LoadFormat::Docx, doc->get_OriginalLoadFormat());
        //ExEnd
    }

    void FootnoteColumns()
    {
        //ExStart
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.Columns
        //ExSummary:Shows how to split the footnote section into a given number of columns.
        auto doc = MakeObject<Document>(MyDir + u"Footnotes and endnotes.docx");
        ASSERT_EQ(0, doc->get_FootnoteOptions()->get_Columns());
        //ExSkip

        doc->get_FootnoteOptions()->set_Columns(2);
        doc->Save(ArtifactsDir + u"Document.FootnoteColumns.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Document.FootnoteColumns.docx");

        ASSERT_EQ(2, doc->get_FirstSection()->get_PageSetup()->get_FootnoteOptions()->get_Columns());
    }

    void Compare()
    {
        //ExStart
        //ExFor:Document.Compare(Document, String, DateTime)
        //ExFor:RevisionCollection.AcceptAll
        //ExSummary:Shows how to compare documents.
        auto docOriginal = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(docOriginal);
        builder->Writeln(u"This is the original document.");

        auto docEdited = MakeObject<Document>();
        builder = MakeObject<DocumentBuilder>(docEdited);
        builder->Writeln(u"This is the edited document.");

        // Comparing documents with revisions will throw an exception.
        if (docOriginal->get_Revisions()->get_Count() == 0 && docEdited->get_Revisions()->get_Count() == 0)
        {
            docOriginal->Compare(docEdited, u"authorName", System::DateTime::get_Now());
        }

        // After the comparison, the original document will gain a new revision
        // for every element that is different in the edited document.
        ASSERT_EQ(2, docOriginal->get_Revisions()->get_Count());
        //ExSkip
        for (const auto& r : System::IterateOver(docOriginal->get_Revisions()))
        {
            std::cout << String::Format(u"Revision type: {0}, on a node of type \"{1}\"", r->get_RevisionType(), r->get_ParentNode()->get_NodeType())
                      << std::endl;
            std::cout << "\tChanged text: \"" << r->get_ParentNode()->GetText() << "\"" << std::endl;
        }

        // Accepting these revisions will transform the original document into the edited document.
        docOriginal->get_Revisions()->AcceptAll();

        ASSERT_EQ(docOriginal->GetText(), docEdited->GetText());
        //ExEnd

        docOriginal = DocumentHelper::SaveOpen(docOriginal);
        ASSERT_EQ(0, docOriginal->get_Revisions()->get_Count());
    }

    void CompareDocumentWithRevisions()
    {
        auto doc1 = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc1);
        builder->Writeln(u"Hello world! This text is not a revision.");

        auto docWithRevision = MakeObject<Document>();
        builder = MakeObject<DocumentBuilder>(docWithRevision);

        docWithRevision->StartTrackRevisions(u"John Doe");
        builder->Writeln(u"This is a revision.");

        ASSERT_THROW(docWithRevision->Compare(doc1, u"John Doe", System::DateTime::get_Now()), System::InvalidOperationException);
    }

    void CompareOptions_()
    {
        //ExStart
        //ExFor:CompareOptions
        //ExFor:CompareOptions.IgnoreFormatting
        //ExFor:CompareOptions.IgnoreCaseChanges
        //ExFor:CompareOptions.IgnoreComments
        //ExFor:CompareOptions.IgnoreTables
        //ExFor:CompareOptions.IgnoreFields
        //ExFor:CompareOptions.IgnoreFootnotes
        //ExFor:CompareOptions.IgnoreTextboxes
        //ExFor:CompareOptions.IgnoreHeadersAndFooters
        //ExFor:CompareOptions.Target
        //ExFor:ComparisonTargetType
        //ExFor:Document.Compare(Document, String, DateTime, CompareOptions)
        //ExSummary:Shows how to filter specific types of document elements when making a comparison.
        // Create the original document and populate it with various kinds of elements.
        auto docOriginal = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(docOriginal);

        // Paragraph text referenced with an endnote:
        builder->Writeln(u"Hello world! This is the first paragraph.");
        builder->InsertFootnote(FootnoteType::Endnote, u"Original endnote text.");

        // Table:
        builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Original cell 1 text");
        builder->InsertCell();
        builder->Write(u"Original cell 2 text");
        builder->EndTable();

        // Textbox:
        SharedPtr<Shape> textBox = builder->InsertShape(ShapeType::TextBox, 150, 20);
        builder->MoveTo(textBox->get_FirstParagraph());
        builder->Write(u"Original textbox contents");

        // DATE field:
        builder->MoveTo(docOriginal->get_FirstSection()->get_Body()->AppendParagraph(u""));
        builder->InsertField(u" DATE ");

        // Comment:
        auto newComment = MakeObject<Comment>(docOriginal, u"John Doe", u"J.D.", System::DateTime::get_Now());
        newComment->SetText(u"Original comment.");
        builder->get_CurrentParagraph()->AppendChild(newComment);

        // Header:
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Writeln(u"Original header contents.");

        // Create a clone of our document and perform a quick edit on each of the cloned document's elements.
        auto docEdited = System::DynamicCast<Document>(docOriginal->Clone(true));
        SharedPtr<Paragraph> firstParagraph = docEdited->get_FirstSection()->get_Body()->get_FirstParagraph();

        firstParagraph->get_Runs()->idx_get(0)->set_Text(u"hello world! this is the first paragraph, after editing.");
        firstParagraph->get_ParagraphFormat()->set_Style(docEdited->get_Styles()->idx_get(StyleIdentifier::Heading1));
        (System::DynamicCast<Footnote>(docEdited->GetChild(NodeType::Footnote, 0, true)))
            ->get_FirstParagraph()
            ->get_Runs()
            ->idx_get(1)
            ->set_Text(u"Edited endnote text.");
        (System::DynamicCast<Table>(docEdited->GetChild(NodeType::Table, 0, true)))
            ->get_FirstRow()
            ->get_Cells()
            ->idx_get(1)
            ->get_FirstParagraph()
            ->get_Runs()
            ->idx_get(0)
            ->set_Text(u"Edited Cell 2 contents");
        (System::DynamicCast<Shape>(docEdited->GetChild(NodeType::Shape, 0, true)))
            ->get_FirstParagraph()
            ->get_Runs()
            ->idx_get(0)
            ->set_Text(u"Edited textbox contents");
        (System::DynamicCast<FieldDate>(docEdited->get_Range()->get_Fields()->idx_get(0)))->set_UseLunarCalendar(true);
        (System::DynamicCast<Comment>(docEdited->GetChild(NodeType::Comment, 0, true)))
            ->get_FirstParagraph()
            ->get_Runs()
            ->idx_get(0)
            ->set_Text(u"Edited comment.");
        docEdited->get_FirstSection()
            ->get_HeadersFooters()
            ->idx_get(HeaderFooterType::HeaderPrimary)
            ->get_FirstParagraph()
            ->get_Runs()
            ->idx_get(0)
            ->set_Text(u"Edited header contents.");

        // Comparing documents creates a revision for every edit in the edited document.
        // A CompareOptions object has a series of flags that can suppress revisions
        // on each respective type of element, effectively ignoring their change.
        auto compareOptions = MakeObject<Aspose::Words::Comparing::CompareOptions>();
        compareOptions->set_IgnoreFormatting(false);
        compareOptions->set_IgnoreCaseChanges(false);
        compareOptions->set_IgnoreComments(false);
        compareOptions->set_IgnoreTables(false);
        compareOptions->set_IgnoreFields(false);
        compareOptions->set_IgnoreFootnotes(false);
        compareOptions->set_IgnoreTextboxes(false);
        compareOptions->set_IgnoreHeadersAndFooters(false);
        compareOptions->set_Target(ComparisonTargetType::New);

        docOriginal->Compare(docEdited, u"John Doe", System::DateTime::get_Now(), compareOptions);
        docOriginal->Save(ArtifactsDir + u"Document.CompareOptions.docx");
        //ExEnd

        docOriginal = MakeObject<Document>(ArtifactsDir + u"Document.CompareOptions.docx");

        TestUtil::VerifyFootnote(FootnoteType::Endnote, true, String::Empty, u"OriginalEdited endnote text.",
                                 System::DynamicCast<Footnote>(docOriginal->GetChild(NodeType::Footnote, 0, true)));
    }

    void IgnoreDmlUniqueId(bool isIgnoreDmlUniqueId)
    {
        //ExStart
        //ExFor:CompareOptions.IgnoreDmlUniqueId
        //ExSummary:Shows how to compare documents ignoring DML unique ID.
        auto docA = MakeObject<Document>(MyDir + u"DML unique ID original.docx");
        auto docB = MakeObject<Document>(MyDir + u"DML unique ID compare.docx");

        // By default, Aspose.Words do not ignore DML's unique ID, and the revisions count was 2.
        // If we are ignoring DML's unique ID, and revisions count were 0.
        auto compareOptions = MakeObject<Aspose::Words::Comparing::CompareOptions>();
        compareOptions->set_IgnoreDmlUniqueId(isIgnoreDmlUniqueId);

        docA->Compare(docB, u"Aspose.Words", System::DateTime::get_Now(), compareOptions);

        ASSERT_EQ(isIgnoreDmlUniqueId ? 0 : 2, docA->get_Revisions()->get_Count());
        //ExEnd
    }

    void RemoveExternalSchemaReferences()
    {
        //ExStart
        //ExFor:Document.RemoveExternalSchemaReferences
        //ExSummary:Shows how to remove all external XML schema references from a document.
        auto doc = MakeObject<Document>(MyDir + u"External XML schema.docx");

        doc->RemoveExternalSchemaReferences();
        //ExEnd
    }

    void TrackRevisions()
    {
        //ExStart
        //ExFor:Document.StartTrackRevisions(String)
        //ExFor:Document.StartTrackRevisions(String, DateTime)
        //ExFor:Document.StopTrackRevisions
        //ExSummary:Shows how to track revisions while editing a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Editing a document usually does not count as a revision until we begin tracking them.
        builder->Write(u"Hello world! ");

        ASSERT_EQ(0, doc->get_Revisions()->get_Count());
        ASSERT_FALSE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0)->get_IsInsertRevision());

        doc->StartTrackRevisions(u"John Doe");

        builder->Write(u"Hello again! ");

        ASSERT_EQ(1, doc->get_Revisions()->get_Count());
        ASSERT_TRUE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(1)->get_IsInsertRevision());
        ASSERT_EQ(u"John Doe", doc->get_Revisions()->idx_get(0)->get_Author());
        ASSERT_LE((System::DateTime::get_Now() - doc->get_Revisions()->idx_get(0)->get_DateTime()).get_Milliseconds(), 10);

        // Stop tracking revisions to not count any future edits as revisions.
        doc->StopTrackRevisions();
        builder->Write(u"Hello again! ");

        ASSERT_EQ(1, doc->get_Revisions()->get_Count());
        ASSERT_FALSE(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(2)->get_IsInsertRevision());

        // Creating revisions gives them a date and time of the operation.
        // We can disable this by passing DateTime.MinValue when we start tracking revisions.
        doc->StartTrackRevisions(u"John Doe", System::DateTime::MinValue);
        builder->Write(u"Hello again! ");

        ASSERT_EQ(2, doc->get_Revisions()->get_Count());
        ASSERT_EQ(u"John Doe", doc->get_Revisions()->idx_get(1)->get_Author());
        ASSERT_EQ(System::DateTime::MinValue, doc->get_Revisions()->idx_get(1)->get_DateTime());

        // We can accept/reject these revisions programmatically
        // by calling methods such as Document.AcceptAllRevisions, or each revision's Accept method.
        // In Microsoft Word, we can process them manually via "Review" -> "Changes".
        doc->Save(ArtifactsDir + u"Document.StartTrackRevisions.docx");
        //ExEnd
    }

    void AcceptAllRevisions()
    {
        //ExStart
        //ExFor:Document.AcceptAllRevisions
        //ExSummary:Shows how to accept all tracking changes in the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Edit the document while tracking changes to create a few revisions.
        doc->StartTrackRevisions(u"John Doe");
        builder->Write(u"Hello world! ");
        builder->Write(u"Hello again! ");
        builder->Write(u"This is another revision.");
        doc->StopTrackRevisions();

        ASSERT_EQ(3, doc->get_Revisions()->get_Count());

        // We can iterate through every revision and accept/reject it as a part of our document.
        // If we know we wish to accept every revision, we can do it more straightforwardly so by calling this method.
        doc->AcceptAllRevisions();

        ASSERT_EQ(0, doc->get_Revisions()->get_Count());
        ASSERT_EQ(u"Hello world! Hello again! This is another revision.", doc->GetText().Trim());
        //ExEnd
    }

    void GetRevisedPropertiesOfList()
    {
        //ExStart
        //ExFor:RevisionsView
        //ExFor:Document.RevisionsView
        //ExSummary:Shows how to switch between the revised and the original view of a document.
        auto doc = MakeObject<Document>(MyDir + u"Revisions at list levels.docx");
        doc->UpdateListLabels();

        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
        ASSERT_EQ(u"1.", paragraphs->idx_get(0)->get_ListLabel()->get_LabelString());
        ASSERT_EQ(u"a.", paragraphs->idx_get(1)->get_ListLabel()->get_LabelString());
        ASSERT_EQ(String::Empty, paragraphs->idx_get(2)->get_ListLabel()->get_LabelString());

        // View the document object as if all the revisions are accepted. Currently supports list labels.
        doc->set_RevisionsView(RevisionsView::Final);

        ASSERT_EQ(String::Empty, paragraphs->idx_get(0)->get_ListLabel()->get_LabelString());
        ASSERT_EQ(u"1.", paragraphs->idx_get(1)->get_ListLabel()->get_LabelString());
        ASSERT_EQ(u"a.", paragraphs->idx_get(2)->get_ListLabel()->get_LabelString());
        //ExEnd

        doc->set_RevisionsView(RevisionsView::Original);
        doc->AcceptAllRevisions();

        ASSERT_EQ(u"a.", paragraphs->idx_get(0)->get_ListLabel()->get_LabelString());
        ASSERT_EQ(String::Empty, paragraphs->idx_get(1)->get_ListLabel()->get_LabelString());
        ASSERT_EQ(u"b.", paragraphs->idx_get(2)->get_ListLabel()->get_LabelString());
    }

    void UpdateThumbnail()
    {
        //ExStart
        //ExFor:Document.UpdateThumbnail()
        //ExFor:Document.UpdateThumbnail(ThumbnailGeneratingOptions)
        //ExFor:ThumbnailGeneratingOptions
        //ExFor:ThumbnailGeneratingOptions.GenerateFromFirstPage
        //ExFor:ThumbnailGeneratingOptions.ThumbnailSize
        //ExSummary:Shows how to update a document's thumbnail.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");
        builder->InsertImage(ImageDir + u"Logo.jpg");

        // There are two ways of setting a thumbnail image when saving a document to .epub.
        // 1 -  Use the document's first page:
        doc->UpdateThumbnail();
        doc->Save(ArtifactsDir + u"Document.UpdateThumbnail.FirstPage.epub");

        // 2 -  Use the first image found in the document:
        auto options = MakeObject<ThumbnailGeneratingOptions>();
        ASPOSE_ASSERT_EQ(System::Drawing::Size(600, 900), options->get_ThumbnailSize());
        //ExSkip
        ASSERT_TRUE(options->get_GenerateFromFirstPage());
        //ExSkip
        options->set_ThumbnailSize(System::Drawing::Size(400, 400));
        options->set_GenerateFromFirstPage(false);

        doc->UpdateThumbnail(options);
        doc->Save(ArtifactsDir + u"Document.UpdateThumbnail.FirstImage.epub");
        //ExEnd
    }

    void HyphenationOptions_()
    {
        //ExStart
        //ExFor:Document.HyphenationOptions
        //ExFor:HyphenationOptions
        //ExFor:HyphenationOptions.AutoHyphenation
        //ExFor:HyphenationOptions.ConsecutiveHyphenLimit
        //ExFor:HyphenationOptions.HyphenationZone
        //ExFor:HyphenationOptions.HyphenateCaps
        //ExSummary:Shows how to configure automatic hyphenation.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Size(24);
        builder->Writeln(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") +
                         u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        doc->get_HyphenationOptions()->set_AutoHyphenation(true);
        doc->get_HyphenationOptions()->set_ConsecutiveHyphenLimit(2);
        doc->get_HyphenationOptions()->set_HyphenationZone(720);
        doc->get_HyphenationOptions()->set_HyphenateCaps(true);

        doc->Save(ArtifactsDir + u"Document.HyphenationOptions.docx");
        //ExEnd

        ASPOSE_ASSERT_EQ(true, doc->get_HyphenationOptions()->get_AutoHyphenation());
        ASSERT_EQ(2, doc->get_HyphenationOptions()->get_ConsecutiveHyphenLimit());
        ASSERT_EQ(720, doc->get_HyphenationOptions()->get_HyphenationZone());
        ASPOSE_ASSERT_EQ(true, doc->get_HyphenationOptions()->get_HyphenateCaps());

        ASSERT_TRUE(DocumentHelper::CompareDocs(ArtifactsDir + u"Document.HyphenationOptions.docx", GoldsDir + u"Document.HyphenationOptions Gold.docx"));
    }

    void HyphenationOptionsDefaultValues()
    {
        auto doc = MakeObject<Document>();
        doc = DocumentHelper::SaveOpen(doc);

        ASPOSE_ASSERT_EQ(false, doc->get_HyphenationOptions()->get_AutoHyphenation());
        ASSERT_EQ(0, doc->get_HyphenationOptions()->get_ConsecutiveHyphenLimit());
        ASSERT_EQ(360, doc->get_HyphenationOptions()->get_HyphenationZone());
        // 0.25 inch
        ASPOSE_ASSERT_EQ(true, doc->get_HyphenationOptions()->get_HyphenateCaps());
    }

    void HyphenationOptionsExceptions()
    {
        auto doc = MakeObject<Document>();

        doc->get_HyphenationOptions()->set_ConsecutiveHyphenLimit(0);
        ASSERT_THROW(doc->get_HyphenationOptions()->set_HyphenationZone(0), System::ArgumentOutOfRangeException);

        ASSERT_THROW(doc->get_HyphenationOptions()->set_ConsecutiveHyphenLimit(-1), System::ArgumentOutOfRangeException);
        doc->get_HyphenationOptions()->set_HyphenationZone(360);
    }

    void OoxmlComplianceVersion()
    {
        //ExStart
        //ExFor:Document.Compliance
        //ExSummary:Shows how to read a loaded document's Open Office XML compliance version.
        // The compliance version varies between documents created by different versions of Microsoft Word.
        auto doc = MakeObject<Document>(MyDir + u"Document.doc");

        ASSERT_EQ(doc->get_Compliance(), OoxmlCompliance::Ecma376_2006);

        doc = MakeObject<Document>(MyDir + u"Document.docx");

        ASSERT_EQ(doc->get_Compliance(), OoxmlCompliance::Iso29500_2008_Transitional);
        //ExEnd
    }

    void ImageSaveOptions_()
    {
        //ExStart
        //ExFor:Document.Save(String, Saving.SaveOptions)
        //ExFor:SaveOptions.UseAntiAliasing
        //ExFor:SaveOptions.UseHighQualityRendering
        //ExSummary:Shows how to improve the quality of a rendered document with SaveOptions.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Size(60);
        builder->Writeln(u"Some text.");

        SharedPtr<SaveOptions> options = MakeObject<Aspose::Words::Saving::ImageSaveOptions>(SaveFormat::Jpeg);
        ASSERT_FALSE(options->get_UseAntiAliasing());
        //ExSkip
        ASSERT_FALSE(options->get_UseHighQualityRendering());
        //ExSkip

        doc->Save(ArtifactsDir + u"Document.ImageSaveOptions.Default.jpg", options);

        options->set_UseAntiAliasing(true);
        options->set_UseHighQualityRendering(true);

        doc->Save(ArtifactsDir + u"Document.ImageSaveOptions.HighQuality.jpg", options);
        //ExEnd

        TestUtil::VerifyImage(794, 1122, ArtifactsDir + u"Document.ImageSaveOptions.Default.jpg");
        TestUtil::VerifyImage(794, 1122, ArtifactsDir + u"Document.ImageSaveOptions.HighQuality.jpg");
    }

    void Cleanup()
    {
        //ExStart
        //ExFor:Document.Cleanup()
        //ExSummary:Shows how to remove unused custom styles from a document.
        auto doc = MakeObject<Document>();

        doc->get_Styles()->Add(StyleType::List, u"MyListStyle1");
        doc->get_Styles()->Add(StyleType::List, u"MyListStyle2");
        doc->get_Styles()->Add(StyleType::Character, u"MyParagraphStyle1");
        doc->get_Styles()->Add(StyleType::Character, u"MyParagraphStyle2");

        // Combined with the built-in styles, the document now has eight styles.
        // A custom style counts as "used" while applied to some part of the document,
        // which means that the four styles we added are currently unused.
        ASSERT_EQ(8, doc->get_Styles()->get_Count());

        // Apply a custom character style, and then a custom list style. Doing so will mark the styles as "used".
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_Font()->set_Style(doc->get_Styles()->idx_get(u"MyParagraphStyle1"));
        builder->Writeln(u"Hello world!");

        SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(doc->get_Styles()->idx_get(u"MyListStyle1"));
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

    void AutomaticallyUpdateStyles()
    {
        //ExStart
        //ExFor:Document.AutomaticallyUpdateStyles
        //ExSummary:Shows how to attach a template to a document.
        auto doc = MakeObject<Document>();

        // Microsoft Word documents by default come with an attached template called "Normal.dotm".
        // There is no default template for blank Aspose.Words documents.
        ASSERT_EQ(String::Empty, doc->get_AttachedTemplate());

        // Attach a template, then set the flag to apply style changes
        // within the template to styles in our document.
        doc->set_AttachedTemplate(MyDir + u"Business brochure.dotx");
        doc->set_AutomaticallyUpdateStyles(true);

        doc->Save(ArtifactsDir + u"Document.AutomaticallyUpdateStyles.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Document.AutomaticallyUpdateStyles.docx");

        ASSERT_TRUE(doc->get_AutomaticallyUpdateStyles());
        ASSERT_EQ(MyDir + u"Business brochure.dotx", doc->get_AttachedTemplate());
        ASSERT_TRUE(System::IO::File::Exists(doc->get_AttachedTemplate()));
    }

    void DefaultTemplate()
    {
        //ExStart
        //ExFor:Document.AttachedTemplate
        //ExFor:Document.AutomaticallyUpdateStyles
        //ExFor:SaveOptions.CreateSaveOptions(String)
        //ExFor:SaveOptions.DefaultTemplate
        //ExSummary:Shows how to set a default template for documents that do not have attached templates.
        auto doc = MakeObject<Document>();

        // Enable automatic style updating, but do not attach a template document.
        doc->set_AutomaticallyUpdateStyles(true);

        ASSERT_EQ(String::Empty, doc->get_AttachedTemplate());

        // Since there is no template document, the document had nowhere to track style changes.
        // Use a SaveOptions object to automatically set a template
        // if a document that we are saving does not have one.
        SharedPtr<SaveOptions> options = SaveOptions::CreateSaveOptions(u"Document.DefaultTemplate.docx");
        options->set_DefaultTemplate(MyDir + u"Business brochure.dotx");

        doc->Save(ArtifactsDir + u"Document.DefaultTemplate.docx", options);
        //ExEnd

        ASSERT_TRUE(System::IO::File::Exists(options->get_DefaultTemplate()));
    }

    void UseSubstitutions()
    {
        //ExStart
        //ExFor:FindReplaceOptions.UseSubstitutions
        //ExFor:FindReplaceOptions.LegacyMode
        //ExSummary:Shows how to recognize and use substitutions within replacement patterns.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Jason gave money to Paul.");

        auto regex = MakeObject<System::Text::RegularExpressions::Regex>(u"([A-z]+) gave money to ([A-z]+)");

        auto options = MakeObject<FindReplaceOptions>();
        options->set_UseSubstitutions(true);

        // Using legacy mode does not support many advanced features, so we need to set it to 'false'.
        options->set_LegacyMode(false);

        doc->get_Range()->Replace(regex, u"$2 took money from $1", options);

        ASSERT_EQ(doc->GetText(), u"Paul took money from Jason.\f");
        //ExEnd
    }

    void SetInvalidateFieldTypes()
    {
        //ExStart
        //ExFor:Document.NormalizeFieldTypes
        //ExSummary:Shows how to get the keep a field's type up to date with its field code.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Field> field = builder->InsertField(u"DATE", nullptr);

        // Aspose.Words automatically detects field types based on field codes.
        ASSERT_EQ(FieldType::FieldDate, field->get_Type());

        // Manually change the raw text of the field, which determines the field code.
        auto fieldText = System::DynamicCast<Run>(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetChildNodes(NodeType::Run, true)->idx_get(0));
        ASSERT_EQ(u"DATE", fieldText->get_Text());
        //ExSkip
        fieldText->set_Text(u"PAGE");

        // Changing the field code has changed this field to one of a different type,
        // but the field's type properties still display the old type.
        ASSERT_EQ(u"PAGE", field->GetFieldCode());
        ASSERT_EQ(FieldType::FieldDate, field->get_Type());
        ASSERT_EQ(FieldType::FieldDate, field->get_Start()->get_FieldType());
        ASSERT_EQ(FieldType::FieldDate, field->get_Separator()->get_FieldType());
        ASSERT_EQ(FieldType::FieldDate, field->get_End()->get_FieldType());

        // Update those properties with this method to display current value.
        doc->NormalizeFieldTypes();

        ASSERT_EQ(FieldType::FieldPage, field->get_Type());
        ASSERT_EQ(FieldType::FieldPage, field->get_Start()->get_FieldType());
        ASSERT_EQ(FieldType::FieldPage, field->get_Separator()->get_FieldType());
        ASSERT_EQ(FieldType::FieldPage, field->get_End()->get_FieldType());
        //ExEnd
    }

    void LayoutOptionsRevisions()
    {
        //ExStart
        //ExFor:Document.LayoutOptions
        //ExFor:LayoutOptions
        //ExFor:LayoutOptions.RevisionOptions
        //ExFor:RevisionColor
        //ExFor:RevisionOptions
        //ExFor:RevisionOptions.InsertedTextColor
        //ExFor:RevisionOptions.ShowRevisionBars
        //ExSummary:Shows how to alter the appearance of revisions in a rendered output document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a revision, then change the color of all revisions to green.
        builder->Writeln(u"This is not a revision.");
        doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
        ASSERT_EQ(RevisionColor::ByAuthor, doc->get_LayoutOptions()->get_RevisionOptions()->get_InsertedTextColor());
        //ExSkip
        ASSERT_TRUE(doc->get_LayoutOptions()->get_RevisionOptions()->get_ShowRevisionBars());
        //ExSkip
        builder->Writeln(u"This is a revision.");
        doc->StopTrackRevisions();
        builder->Writeln(u"This is not a revision.");

        // Remove the bar that appears to the left of every revised line.
        doc->get_LayoutOptions()->get_RevisionOptions()->set_InsertedTextColor(RevisionColor::BrightGreen);
        doc->get_LayoutOptions()->get_RevisionOptions()->set_ShowRevisionBars(false);

        doc->Save(ArtifactsDir + u"Document.LayoutOptionsRevisions.pdf");
        //ExEnd
    }

    void LayoutOptionsHiddenText(bool showHiddenText)
    {
        //ExStart
        //ExFor:Document.LayoutOptions
        //ExFor:LayoutOptions
        //ExFor:Layout.LayoutOptions.ShowHiddenText
        //ExSummary:Shows how to hide text in a rendered output document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        ASSERT_FALSE(doc->get_LayoutOptions()->get_ShowHiddenText());
        //ExSkip

        // Insert hidden text, then specify whether we wish to omit it from a rendered document.
        builder->Writeln(u"This text is not hidden.");
        builder->get_Font()->set_Hidden(true);
        builder->Writeln(u"This text is hidden.");

        doc->get_LayoutOptions()->set_ShowHiddenText(showHiddenText);

        doc->Save(ArtifactsDir + u"Document.LayoutOptionsHiddenText.pdf");
        //ExEnd
    }

    void LayoutOptionsParagraphMarks(bool showParagraphMarks)
    {
        //ExStart
        //ExFor:Document.LayoutOptions
        //ExFor:LayoutOptions
        //ExFor:Layout.LayoutOptions.ShowParagraphMarks
        //ExSummary:Shows how to show paragraph marks in a rendered output document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        ASSERT_FALSE(doc->get_LayoutOptions()->get_ShowParagraphMarks());
        //ExSkip

        // Add some paragraphs, then enable paragraph marks to show the ends of paragraphs
        // with a pilcrow (¶) symbol when we render the document.
        builder->Writeln(u"Hello world!");
        builder->Writeln(u"Hello again!");

        doc->get_LayoutOptions()->set_ShowParagraphMarks(showParagraphMarks);

        doc->Save(ArtifactsDir + u"Document.LayoutOptionsParagraphMarks.pdf");
        //ExEnd
    }

    void UpdatePageLayout()
    {
        //ExStart
        //ExFor:StyleCollection.Item(String)
        //ExFor:SectionCollection.Item(Int32)
        //ExFor:Document.UpdatePageLayout
        //ExSummary:Shows when to recalculate the page layout of the document.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Saving a document to PDF, to an image, or printing for the first time will automatically
        // cache the layout of the document within its pages.
        doc->Save(ArtifactsDir + u"Document.UpdatePageLayout.1.pdf");

        // Modify the document in some way.
        doc->get_Styles()->idx_get(u"Normal")->get_Font()->set_Size(6);
        doc->get_Sections()->idx_get(0)->get_PageSetup()->set_Orientation(Orientation::Landscape);

        // In the current version of Aspose.Words, modifying the document does not automatically rebuild
        // the cached page layout. If we wish for the cached layout
        // to stay up to date, we will need to update it manually.
        doc->UpdatePageLayout();

        doc->Save(ArtifactsDir + u"Document.UpdatePageLayout.2.pdf");
        //ExEnd
    }

    void DocPackageCustomParts()
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
        auto doc = MakeObject<Document>(MyDir + u"Custom parts OOXML package.docx");

        ASSERT_EQ(2, doc->get_PackageCustomParts()->get_Count());

        // Clone the second part, then add the clone to the collection.
        SharedPtr<CustomPart> clonedPart = doc->get_PackageCustomParts()->idx_get(1)->Clone();
        doc->get_PackageCustomParts()->Add(clonedPart);
        TestDocPackageCustomParts(doc->get_PackageCustomParts());
        //ExSkip

        ASSERT_EQ(3, doc->get_PackageCustomParts()->get_Count());

        // Enumerate over the collection and print every part.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<CustomPart>>> enumerator = doc->get_PackageCustomParts()->GetEnumerator();
            int index = 0;
            while (enumerator->MoveNext())
            {
                std::cout << "Part index " << index << ":" << std::endl;
                std::cout << "\tName:\t\t\t\t" << enumerator->get_Current()->get_Name() << std::endl;
                std::cout << "\tContent type:\t\t" << enumerator->get_Current()->get_ContentType() << std::endl;
                std::cout << "\tRelationship type:\t" << enumerator->get_Current()->get_RelationshipType() << std::endl;
                std::cout << (enumerator->get_Current()->get_IsExternal()
                                  ? u"\tSourced from outside the document"
                                  : String::Format(u"\tStored within the document, length: {0} bytes", enumerator->get_Current()->get_Data()->get_Length()))
                          << std::endl;
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

    static void TestDocPackageCustomParts(SharedPtr<CustomPartCollection> parts)
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

    void ShadeFormData(bool useGreyShading)
    {
        //ExStart
        //ExFor:Document.ShadeFormData
        //ExSummary:Shows how to apply gray shading to form fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        ASSERT_TRUE(doc->get_ShadeFormData());
        //ExSkip

        builder->Write(u"Hello world! ");
        builder->InsertTextInput(u"My form field", TextFormFieldType::Regular, u"", u"Text contents of form field, which are shaded in grey by default.", 0);

        // We can turn the grey shading off, so the bookmarked text will blend in with the other text.
        doc->set_ShadeFormData(useGreyShading);
        doc->Save(ArtifactsDir + u"Document.ShadeFormData.docx");
        //ExEnd
    }

    void VersionsCount()
    {
        //ExStart
        //ExFor:Document.VersionsCount
        //ExSummary:Shows how to work with the versions count feature of older Microsoft Word documents.
        auto doc = MakeObject<Document>(MyDir + u"Versions.doc");

        // We can read this property of a document, but we cannot preserve it while saving.
        ASSERT_EQ(4, doc->get_VersionsCount());

        doc->Save(ArtifactsDir + u"Document.VersionsCount.doc");
        doc = MakeObject<Document>(ArtifactsDir + u"Document.VersionsCount.doc");

        ASSERT_EQ(0, doc->get_VersionsCount());
        //ExEnd
    }

    void WriteProtection_()
    {
        //ExStart
        //ExFor:Document.WriteProtection
        //ExFor:WriteProtection
        //ExFor:WriteProtection.IsWriteProtected
        //ExFor:WriteProtection.ReadOnlyRecommended
        //ExFor:WriteProtection.SetPassword(String)
        //ExFor:WriteProtection.ValidatePassword(String)
        //ExSummary:Shows how to protect a document with a password.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
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
        doc->Save(ArtifactsDir + u"Document.WriteProtection.docx");
        doc = MakeObject<Document>(ArtifactsDir + u"Document.WriteProtection.docx");

        ASSERT_TRUE(doc->get_WriteProtection()->get_IsWriteProtected());

        builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveToDocumentEnd();
        builder->Writeln(u"Writing text in a protected document.");

        ASSERT_EQ(String(u"Hello world! This document is protected.") + u"\rWriting text in a protected document.", doc->GetText().Trim());
        //ExEnd
        ASSERT_TRUE(doc->get_WriteProtection()->get_ReadOnlyRecommended());
        ASSERT_TRUE(doc->get_WriteProtection()->ValidatePassword(u"MyPassword"));
        ASSERT_FALSE(doc->get_WriteProtection()->ValidatePassword(u"wrongpassword"));
    }

    void RemovePersonalInformation(bool saveWithoutPersonalInfo)
    {
        //ExStart
        //ExFor:Document.RemovePersonalInformation
        //ExSummary:Shows how to enable the removal of personal information during a manual save.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

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
        doc->Save(ArtifactsDir + u"Document.RemovePersonalInformation.docx");
        doc = MakeObject<Document>(ArtifactsDir + u"Document.RemovePersonalInformation.docx");

        ASPOSE_ASSERT_EQ(saveWithoutPersonalInfo, doc->get_RemovePersonalInformation());
        ASSERT_EQ(u"John Doe", doc->get_BuiltInDocumentProperties()->get_Author());
        ASSERT_EQ(u"Placeholder Inc.", doc->get_BuiltInDocumentProperties()->get_Company());
        ASSERT_EQ(u"John Doe", doc->get_Revisions()->idx_get(0)->get_Author());
        //ExEnd
    }

    void ShowComments()
    {
        //ExStart
        //ExFor:LayoutOptions.CommentDisplayMode
        //ExFor:CommentDisplayMode
        //ExSummary:Shows how to show comments when saving a document to a rendered format.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Hello world!");

        auto comment = MakeObject<Comment>(doc, u"John Doe", u"J.D.", System::DateTime::get_Now());
        comment->SetText(u"My comment.");
        builder->get_CurrentParagraph()->AppendChild(comment);

        // ShowInAnnotations is only available in Pdf1.7 and Pdf1.5 formats.
        // In other formats, it will work similarly to Hide.
        doc->get_LayoutOptions()->set_CommentDisplayMode(CommentDisplayMode::ShowInAnnotations);

        doc->Save(ArtifactsDir + u"Document.ShowCommentsInAnnotations.pdf");

        // Note that it's required to rebuild the document page layout (via Document.UpdatePageLayout() method)
        // after changing the Document.LayoutOptions values.
        doc->get_LayoutOptions()->set_CommentDisplayMode(CommentDisplayMode::ShowInBalloons);
        doc->UpdatePageLayout();

        doc->Save(ArtifactsDir + u"Document.ShowCommentsInBalloons.pdf");
        //ExEnd
    }

    void CopyTemplateStylesViaDocument()
    {
        //ExStart
        //ExFor:Document.CopyStylesFromTemplate(Document)
        //ExSummary:Shows how to copies styles from the template to a document via Document.
        auto template_ = MakeObject<Document>(MyDir + u"Rendering.docx");
        auto target = MakeObject<Document>(MyDir + u"Document.docx");

        ASSERT_EQ(18, template_->get_Styles()->get_Count());
        //ExSkip
        ASSERT_EQ(12, target->get_Styles()->get_Count());
        //ExSkip

        target->CopyStylesFromTemplate(template_);
        ASSERT_EQ(22, target->get_Styles()->get_Count());
        //ExSkip

        //ExEnd
    }

    void CopyTemplateStylesViaDocumentNew()
    {
        //ExStart
        //ExFor:Document.CopyStylesFromTemplate(Document)
        //ExFor:Document.CopyStylesFromTemplate(String)
        //ExSummary:Shows how to copy styles from one document to another.
        // Create a document, and then add styles that we will copy to another document.
        auto template_ = MakeObject<Document>();

        SharedPtr<Style> style = template_->get_Styles()->Add(StyleType::Paragraph, u"TemplateStyle1");
        style->get_Font()->set_Name(u"Times New Roman");
        style->get_Font()->set_Color(System::Drawing::Color::get_Navy());

        style = template_->get_Styles()->Add(StyleType::Paragraph, u"TemplateStyle2");
        style->get_Font()->set_Name(u"Arial");
        style->get_Font()->set_Color(System::Drawing::Color::get_DeepSkyBlue());

        style = template_->get_Styles()->Add(StyleType::Paragraph, u"TemplateStyle3");
        style->get_Font()->set_Name(u"Courier New");
        style->get_Font()->set_Color(System::Drawing::Color::get_RoyalBlue());

        ASSERT_EQ(7, template_->get_Styles()->get_Count());

        // Create a document which we will copy the styles to.
        auto target = MakeObject<Document>();

        // Create a style with the same name as a style from the template document and add it to the target document.
        style = target->get_Styles()->Add(StyleType::Paragraph, u"TemplateStyle3");
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
        target->CopyStylesFromTemplate(MyDir + u"Rendering.docx");

        ASSERT_EQ(21, target->get_Styles()->get_Count());
        //ExEnd
    }

    void ReadMacrosFromExistingDocument()
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
        auto doc = MakeObject<Document>(MyDir + u"VBA project.docm");

        // A VBA project contains a collection of VBA modules.
        SharedPtr<VbaProject> vbaProject = doc->get_VbaProject();
        ASSERT_TRUE(vbaProject->get_IsSigned());
        //ExSkip
        std::cout << (vbaProject->get_IsSigned() ? String::Format(u"Project name: {0} signed; Project code page: {1}; Modules count: {2}\n",
                                                                  vbaProject->get_Name(), vbaProject->get_CodePage(), vbaProject->get_Modules()->LINQ_Count())
                                                 : String::Format(u"Project name: {0} not signed; Project code page: {1}; Modules count: {2}\n",
                                                                  vbaProject->get_Name(), vbaProject->get_CodePage(), vbaProject->get_Modules()->LINQ_Count()))
                  << std::endl;

        SharedPtr<VbaModuleCollection> vbaModules = doc->get_VbaProject()->get_Modules();

        ASSERT_EQ(vbaModules->LINQ_Count(), 3);

        for (const auto& module_ : vbaModules)
        {
            std::cout << "Module name: " << module_->get_Name() << ";\nModule code:\n" << module_->get_SourceCode() << "\n" << std::endl;
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

    void SaveOutputParameters_()
    {
        //ExStart
        //ExFor:SaveOutputParameters
        //ExFor:SaveOutputParameters.ContentType
        //ExSummary:Shows how to access output parameters of a document's save operation.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // After we save a document, we can access the Internet Media Type (MIME type) of the newly created output document.
        SharedPtr<SaveOutputParameters> parameters = doc->Save(ArtifactsDir + u"Document.SaveOutputParameters.doc");

        ASSERT_EQ(u"application/msword", parameters->get_ContentType());

        // This property changes depending on the save format.
        parameters = doc->Save(ArtifactsDir + u"Document.SaveOutputParameters.pdf");

        ASSERT_EQ(u"application/pdf", parameters->get_ContentType());
        //ExEnd
    }

    void SubDocument_()
    {
        //ExStart
        //ExFor:SubDocument
        //ExFor:SubDocument.NodeType
        //ExSummary:Shows how to access a master document's subdocument.
        auto doc = MakeObject<Document>(MyDir + u"Master document.docx");

        SharedPtr<NodeCollection> subDocuments = doc->GetChildNodes(NodeType::SubDocument, true);
        ASSERT_EQ(1, subDocuments->get_Count());
        //ExSkip

        // This node serves as a reference to an external document, and its contents cannot be accessed.
        auto subDocument = System::DynamicCast<SubDocument>(subDocuments->idx_get(0));

        ASSERT_FALSE(subDocument->get_IsComposite());
        //ExEnd
    }

    void CreateWebExtension()
    {
        //ExStart
        //ExFor:BaseWebExtensionCollection`1.Add(`0)
        //ExFor:BaseWebExtensionCollection`1.Clear
        //ExFor:TaskPane
        //ExFor:TaskPane.DockState
        //ExFor:TaskPane.IsVisible
        //ExFor:TaskPane.Width
        //ExFor:TaskPane.IsLocked
        //ExFor:TaskPane.WebExtension
        //ExFor:TaskPane.Row
        //ExFor:WebExtension
        //ExFor:WebExtension.Reference
        //ExFor:WebExtension.Properties
        //ExFor:WebExtension.Bindings
        //ExFor:WebExtension.IsFrozen
        //ExFor:WebExtensionReference.Id
        //ExFor:WebExtensionReference.Version
        //ExFor:WebExtensionReference.StoreType
        //ExFor:WebExtensionReference.Store
        //ExFor:WebExtensionPropertyCollection
        //ExFor:WebExtensionBindingCollection
        //ExFor:WebExtensionProperty.#ctor(String, String)
        //ExFor:WebExtensionBinding.#ctor(String, WebExtensionBindingType, String)
        //ExFor:WebExtensionStoreType
        //ExFor:WebExtensionBindingType
        //ExFor:TaskPaneDockState
        //ExFor:TaskPaneCollection
        //ExSummary:Shows how to add a web extension to a document.
        auto doc = MakeObject<Document>();

        // Create task pane with "MyScript" add-in, which will be used by the document,
        // then set its default location.
        auto myScriptTaskPane = MakeObject<TaskPane>();
        doc->get_WebExtensionTaskPanes()->Add(myScriptTaskPane);
        myScriptTaskPane->set_DockState(TaskPaneDockState::Right);
        myScriptTaskPane->set_IsVisible(true);
        myScriptTaskPane->set_Width(300);
        myScriptTaskPane->set_IsLocked(true);

        // If there are multiple task panes in the same docking location, we can set this index to arrange them.
        myScriptTaskPane->set_Row(1);

        // Create an add-in called "MyScript Math Sample", which the task pane will display within.
        SharedPtr<WebExtension> webExtension = myScriptTaskPane->get_WebExtension();

        // Set application store reference parameters for our add-in, such as the ID.
        webExtension->get_Reference()->set_Id(u"WA104380646");
        webExtension->get_Reference()->set_Version(u"1.0.0.0");
        webExtension->get_Reference()->set_StoreType(WebExtensionStoreType::OMEX);
        webExtension->get_Reference()->set_Store(System::Globalization::CultureInfo::get_CurrentCulture()->get_Name());
        webExtension->get_Properties()->Add(MakeObject<WebExtensionProperty>(u"MyScript", u"MyScript Math Sample"));
        webExtension->get_Bindings()->Add(MakeObject<WebExtensionBinding>(u"MyScript", WebExtensionBindingType::Text, u"104380646"));

        // Allow the user to interact with the add-in.
        webExtension->set_IsFrozen(false);

        // We can access the web extension in Microsoft Word via Developer -> Add-ins.
        doc->Save(ArtifactsDir + u"Document.WebExtension.docx");

        // Remove all web extension task panes at once like this.
        doc->get_WebExtensionTaskPanes()->Clear();

        ASSERT_EQ(0, doc->get_WebExtensionTaskPanes()->get_Count());
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Document.WebExtension.docx");
        myScriptTaskPane = doc->get_WebExtensionTaskPanes()->idx_get(0);

        ASSERT_EQ(TaskPaneDockState::Right, myScriptTaskPane->get_DockState());
        ASSERT_TRUE(myScriptTaskPane->get_IsVisible());
        ASPOSE_ASSERT_EQ(300.0, myScriptTaskPane->get_Width());
        ASSERT_TRUE(myScriptTaskPane->get_IsLocked());
        ASSERT_EQ(1, myScriptTaskPane->get_Row());
        webExtension = myScriptTaskPane->get_WebExtension();

        ASSERT_EQ(u"WA104380646", webExtension->get_Reference()->get_Id());
        ASSERT_EQ(u"1.0.0.0", webExtension->get_Reference()->get_Version());
        ASSERT_EQ(WebExtensionStoreType::OMEX, webExtension->get_Reference()->get_StoreType());
        ASSERT_EQ(System::Globalization::CultureInfo::get_CurrentCulture()->get_Name(), webExtension->get_Reference()->get_Store());

        ASSERT_EQ(u"MyScript", webExtension->get_Properties()->idx_get(0)->get_Name());
        ASSERT_EQ(u"MyScript Math Sample", webExtension->get_Properties()->idx_get(0)->get_Value());

        ASSERT_EQ(u"MyScript", webExtension->get_Bindings()->idx_get(0)->get_Id());
        ASSERT_EQ(WebExtensionBindingType::Text, webExtension->get_Bindings()->idx_get(0)->get_BindingType());
        ASSERT_EQ(u"104380646", webExtension->get_Bindings()->idx_get(0)->get_AppRef());

        ASSERT_FALSE(webExtension->get_IsFrozen());
    }

    void GetWebExtensionInfo()
    {
        //ExStart
        //ExFor:BaseWebExtensionCollection`1
        //ExFor:BaseWebExtensionCollection`1.GetEnumerator
        //ExFor:BaseWebExtensionCollection`1.Remove(Int32)
        //ExFor:BaseWebExtensionCollection`1.Count
        //ExFor:BaseWebExtensionCollection`1.Item(Int32)
        //ExSummary:Shows how to work with a document's collection of web extensions.
        auto doc = MakeObject<Document>(MyDir + u"Web extension.docx");

        ASSERT_EQ(1, doc->get_WebExtensionTaskPanes()->get_Count());

        // Print all properties of the document's web extension.
        SharedPtr<WebExtensionPropertyCollection> webExtensionPropertyCollection =
            doc->get_WebExtensionTaskPanes()->idx_get(0)->get_WebExtension()->get_Properties();
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<WebExtensionProperty>>> enumerator = webExtensionPropertyCollection->GetEnumerator();
            while (enumerator->MoveNext())
            {
                SharedPtr<WebExtensionProperty> webExtensionProperty = enumerator->get_Current();
                std::cout << "Binding name: " << webExtensionProperty->get_Name() << "; Binding value: " << webExtensionProperty->get_Value() << std::endl;
            }
        }

        // Remove the web extension.
        doc->get_WebExtensionTaskPanes()->Remove(0);

        ASSERT_EQ(0, doc->get_WebExtensionTaskPanes()->get_Count());
        //ExEnd
    }

    void EpubCover()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // When saving to .epub, some Microsoft Word document properties convert to .epub metadata.
        doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
        doc->get_BuiltInDocumentProperties()->set_Title(u"My Book Title");

        // The thumbnail we specify here can become the cover image.
        ArrayPtr<uint8_t> image = System::IO::File::ReadAllBytes(ImageDir + u"Transparent background logo.png");
        doc->get_BuiltInDocumentProperties()->set_Thumbnail(image);

        doc->Save(ArtifactsDir + u"Document.EpubCover.epub");
    }

    void TextWatermark()
    {
        //ExStart
        //ExFor:Watermark.SetText(String)
        //ExFor:Watermark.SetText(String, TextWatermarkOptions)
        //ExFor:Watermark.Remove
        //ExFor:TextWatermarkOptions.FontFamily
        //ExFor:TextWatermarkOptions.FontSize
        //ExFor:TextWatermarkOptions.Color
        //ExFor:TextWatermarkOptions.Layout
        //ExFor:TextWatermarkOptions.IsSemitrasparent
        //ExFor:WatermarkLayout
        //ExFor:WatermarkType
        //ExSummary:Shows how to create a text watermark.
        auto doc = MakeObject<Document>();

        // Add a plain text watermark.
        doc->get_Watermark()->SetText(u"Aspose Watermark");

        // If we wish to edit the text formatting using it as a watermark,
        // we can do so by passing a TextWatermarkOptions object when creating the watermark.
        auto textWatermarkOptions = MakeObject<TextWatermarkOptions>();
        textWatermarkOptions->set_FontFamily(u"Arial");
        textWatermarkOptions->set_FontSize(36.0f);
        textWatermarkOptions->set_Color(System::Drawing::Color::get_Black());
        textWatermarkOptions->set_Layout(WatermarkLayout::Diagonal);
        textWatermarkOptions->set_IsSemitrasparent(false);

        doc->get_Watermark()->SetText(u"Aspose Watermark", textWatermarkOptions);

        doc->Save(ArtifactsDir + u"Document.TextWatermark.docx");

        // We can remove a watermark from a document like this.
        if (doc->get_Watermark()->get_Type() == WatermarkType::Text)
        {
            doc->get_Watermark()->Remove();
        }
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Document.TextWatermark.docx");

        ASSERT_EQ(WatermarkType::Text, doc->get_Watermark()->get_Type());
    }

    void ImageWatermark()
    {
        //ExStart
        //ExFor:Watermark.SetImage(Image, ImageWatermarkOptions)
        //ExFor:ImageWatermarkOptions.Scale
        //ExFor:ImageWatermarkOptions.IsWashout
        //ExSummary:Shows how to create a watermark from an image in the local file system.
        auto doc = MakeObject<Document>();

        // Modify the image watermark's appearance with an ImageWatermarkOptions object,
        // then pass it while creating a watermark from an image file.
        auto imageWatermarkOptions = MakeObject<ImageWatermarkOptions>();
        imageWatermarkOptions->set_Scale(5);
        imageWatermarkOptions->set_IsWashout(false);

        doc->get_Watermark()->SetImage(System::Drawing::Image::FromFile(ImageDir + u"Logo.jpg"), imageWatermarkOptions);

        doc->Save(ArtifactsDir + u"Document.ImageWatermark.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Document.ImageWatermark.docx");

        ASSERT_EQ(WatermarkType::Image, doc->get_Watermark()->get_Type());
    }

    void SpellingAndGrammarErrors(bool showErrors)
    {
        //ExStart
        //ExFor:Document.ShowGrammaticalErrors
        //ExFor:Document.ShowSpellingErrors
        //ExSummary:Shows how to show/hide errors in the document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert two sentences with mistakes that would be picked up
        // by the spelling and grammar checkers in Microsoft Word.
        builder->Writeln(u"There is a speling error in this sentence.");
        builder->Writeln(u"Their is a grammatical error in this sentence.");

        // If these options are enabled, then spelling errors will be underlined
        // in the output document by a jagged red line, and a double blue line will highlight grammatical mistakes.
        doc->set_ShowGrammaticalErrors(showErrors);
        doc->set_ShowSpellingErrors(showErrors);

        doc->Save(ArtifactsDir + u"Document.SpellingAndGrammarErrors.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Document.SpellingAndGrammarErrors.docx");

        ASPOSE_ASSERT_EQ(showErrors, doc->get_ShowGrammaticalErrors());
        ASPOSE_ASSERT_EQ(showErrors, doc->get_ShowSpellingErrors());
    }

    void GranularityCompareOption(Granularity granularity)
    {
        //ExStart
        //ExFor:CompareOptions.Granularity
        //ExFor:Granularity
        //ExSummary:Shows to specify a granularity while comparing documents.
        auto docA = MakeObject<Document>();
        auto builderA = MakeObject<DocumentBuilder>(docA);
        builderA->Writeln(u"Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit");

        auto docB = MakeObject<Document>();
        auto builderB = MakeObject<DocumentBuilder>(docB);
        builderB->Writeln(u"Lorems ipsum dolor sit amet consectetur - \"adipiscing\" elit");

        // Specify whether changes are tracking
        // by character ('Granularity.CharLevel'), or by word ('Granularity.WordLevel').
        auto compareOptions = MakeObject<Aspose::Words::Comparing::CompareOptions>();
        compareOptions->set_Granularity(granularity);

        docA->Compare(docB, u"author", System::DateTime::get_Now(), compareOptions);

        // The first document's collection of revision groups contains all the differences between documents.
        SharedPtr<RevisionGroupCollection> groups = docA->get_Revisions()->get_Groups();
        ASSERT_EQ(5, groups->get_Count());
        //ExEnd

        if (granularity == Granularity::CharLevel)
        {
            ASSERT_EQ(RevisionType::Deletion, groups->idx_get(0)->get_RevisionType());
            ASSERT_EQ(u"Alpha ", groups->idx_get(0)->get_Text());

            ASSERT_EQ(RevisionType::Deletion, groups->idx_get(1)->get_RevisionType());
            ASSERT_EQ(u",", groups->idx_get(1)->get_Text());

            ASSERT_EQ(RevisionType::Insertion, groups->idx_get(2)->get_RevisionType());
            ASSERT_EQ(u"s", groups->idx_get(2)->get_Text());

            ASSERT_EQ(RevisionType::Insertion, groups->idx_get(3)->get_RevisionType());
            ASSERT_EQ(u"- \"", groups->idx_get(3)->get_Text());

            ASSERT_EQ(RevisionType::Insertion, groups->idx_get(4)->get_RevisionType());
            ASSERT_EQ(u"\"", groups->idx_get(4)->get_Text());
        }
        else
        {
            ASSERT_EQ(RevisionType::Deletion, groups->idx_get(0)->get_RevisionType());
            ASSERT_EQ(u"Alpha Lorem", groups->idx_get(0)->get_Text());

            ASSERT_EQ(RevisionType::Deletion, groups->idx_get(1)->get_RevisionType());
            ASSERT_EQ(u",", groups->idx_get(1)->get_Text());

            ASSERT_EQ(RevisionType::Insertion, groups->idx_get(2)->get_RevisionType());
            ASSERT_EQ(u"Lorems", groups->idx_get(2)->get_Text());

            ASSERT_EQ(RevisionType::Insertion, groups->idx_get(3)->get_RevisionType());
            ASSERT_EQ(u"- \"", groups->idx_get(3)->get_Text());

            ASSERT_EQ(RevisionType::Insertion, groups->idx_get(4)->get_RevisionType());
            ASSERT_EQ(u"\"", groups->idx_get(4)->get_Text());
        }
    }

    void IgnorePrinterMetrics()
    {
        //ExStart
        //ExFor:LayoutOptions.IgnorePrinterMetrics
        //ExSummary:Shows how to ignore 'Use printer metrics to lay out document' option.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        doc->get_LayoutOptions()->set_IgnorePrinterMetrics(false);

        doc->Save(ArtifactsDir + u"Document.IgnorePrinterMetrics.docx");
        //ExEnd
    }

    void ExtractPages()
    {
        //ExStart
        //ExFor:Document.ExtractPages
        //ExSummary:Shows how to get specified range of pages from the document.
        auto doc = MakeObject<Document>(MyDir + u"Layout entities.docx");

        doc = doc->ExtractPages(0, 2);

        doc->Save(ArtifactsDir + u"Document.ExtractPages.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Document.ExtractPages.docx");
        ASSERT_EQ(doc->get_PageCount(), 2);
    }

    void SpellingOrGrammar(bool checkSpellingGrammar)
    {
        //ExStart
        //ExFor:Document.SpellingChecked
        //ExFor:Document.GrammarChecked
        //ExSummary:Shows how to set spelling or grammar verifying.
        auto doc = MakeObject<Document>();

        // The string with spelling errors.
        doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->Add(MakeObject<Run>(doc, u"The speeling in this documentz is all broked."));

        // Spelling/Grammar check start if we set properties to false.
        // We can see all errors in Microsoft Word via Review -> Spelling & Grammar.
        // Note that Microsoft Word does not start grammar/spell check automatically for DOC and RTF document format.
        doc->set_SpellingChecked(checkSpellingGrammar);
        doc->set_GrammarChecked(checkSpellingGrammar);

        doc->Save(ArtifactsDir + u"Document.SpellingOrGrammar.docx");
        //ExEnd
    }

    void AllowEmbeddingPostScriptFonts()
    {
        //ExStart
        //ExFor:SaveOptions.AllowEmbeddingPostScriptFonts
        //ExSummary:Shows how to save the document with PostScript font.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"PostScriptFont");
        builder->Writeln(u"Some text with PostScript font.");

        // Load the font with PostScript to use in the document.
        auto otf = MakeObject<MemoryFontSource>(System::IO::File::ReadAllBytes(FontsDir + u"AllegroOpen.otf"));
        doc->set_FontSettings(MakeObject<FontSettings>());
        doc->get_FontSettings()->SetFontsSources(MakeArray<SharedPtr<FontSourceBase>>({otf}));

        // Embed TrueType fonts.
        doc->get_FontInfos()->set_EmbedTrueTypeFonts(true);

        // Allow embedding PostScript fonts while embedding TrueType fonts.
        // Microsoft Word does not embed PostScript fonts, but can open documents with embedded fonts of this type.
        SharedPtr<SaveOptions> saveOptions = SaveOptions::CreateSaveOptions(SaveFormat::Docx);
        saveOptions->set_AllowEmbeddingPostScriptFonts(true);

        doc->Save(ArtifactsDir + u"Document.AllowEmbeddingPostScriptFonts.docx", saveOptions);
        //ExEnd
    }

    void Frameset()
    {
        //ExStart
        //ExFor:Document.Frameset
        //ExFor:Frameset
        //ExFor:Frameset.FrameDefaultUrl
        //ExFor:Frameset.IsFrameLinkToFile
        //ExFor:Frameset.ChildFramesets
        //ExSummary:Shows how to access frames on-page.
        // Document contains several frames with links to other documents.
        auto doc = MakeObject<Document>(MyDir + u"Frameset.docx");

        // We can check the default URL (a web page URL or local document) or if the frame is an external resource.
        ASSERT_EQ(u"https://file-examples-com.github.io/uploads/2017/02/file-sample_100kB.docx",
                  doc->get_Frameset()->get_ChildFramesets()->idx_get(0)->get_ChildFramesets()->idx_get(0)->get_FrameDefaultUrl());
        ASSERT_TRUE(doc->get_Frameset()->get_ChildFramesets()->idx_get(0)->get_ChildFramesets()->idx_get(0)->get_IsFrameLinkToFile());

        ASSERT_EQ(u"Document.docx", doc->get_Frameset()->get_ChildFramesets()->idx_get(1)->get_FrameDefaultUrl());
        ASSERT_FALSE(doc->get_Frameset()->get_ChildFramesets()->idx_get(1)->get_IsFrameLinkToFile());

        // Change properties for one of our frames.
        doc->get_Frameset()->get_ChildFramesets()->idx_get(0)->get_ChildFramesets()->idx_get(0)->set_FrameDefaultUrl(
            u"https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx");
        doc->get_Frameset()->get_ChildFramesets()->idx_get(0)->get_ChildFramesets()->idx_get(0)->set_IsFrameLinkToFile(false);
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);

        ASSERT_EQ(u"https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx",
                  doc->get_Frameset()->get_ChildFramesets()->idx_get(0)->get_ChildFramesets()->idx_get(0)->get_FrameDefaultUrl());
        ASSERT_FALSE(doc->get_Frameset()->get_ChildFramesets()->idx_get(0)->get_ChildFramesets()->idx_get(0)->get_IsFrameLinkToFile());
    }
};

} // namespace ApiExamples
