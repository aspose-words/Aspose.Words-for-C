#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <functional>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeMarkupLanguage.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Saving/CompressionLevel.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/exceptions.h>
#include <system/io/file.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"
#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace ApiExamples {

class ExOoxmlSaveOptions : public ApiExampleBase
{
public:
    void Password()
    {
        //ExStart
        //ExFor:OoxmlSaveOptions.Password
        //ExSummary:Shows how to create a password protected Office Open XML document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // Create a SaveOptions object with a password and save our document with it
        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_Password(u"MyPassword");

        doc->Save(ArtifactsDir + u"OoxmlSaveOptions.Password.docx", saveOptions);

        // This document cannot be opened like a normal document

        auto openDocument = [&doc]()
        {
            doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.Password.docx");
        };

        ASSERT_THROW(openDocument(), IncorrectPasswordException);

        // We can open the document and access its contents by passing the correct password to a LoadOptions object
        doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.Password.docx", MakeObject<LoadOptions>(u"MyPassword"));

        ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
        //ExEnd
    }

    void Iso29500Strict()
    {
        //ExStart
        //ExFor:OoxmlSaveOptions
        //ExFor:OoxmlSaveOptions.#ctor()
        //ExFor:OoxmlSaveOptions.SaveFormat
        //ExFor:OoxmlCompliance
        //ExFor:OoxmlSaveOptions.Compliance
        //ExFor:ShapeMarkupLanguage
        //ExSummary:Shows conversion VML shapes to DML using ISO/IEC 29500:2008 Strict compliance level.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set Word2003 version for document, for inserting image as VML shape
        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2003);
        builder->InsertImage(ImageDir + u"Transparent background logo.png");

        ASSERT_EQ(ShapeMarkupLanguage::Vml, (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_MarkupLanguage());

        // Iso29500_2008 does not allow VML shapes
        // You need to use OoxmlCompliance.Iso29500_2008_Strict for converting VML to DML shapes
        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_Compliance(OoxmlCompliance::Iso29500_2008_Strict);
        saveOptions->set_SaveFormat(SaveFormat::Docx);

        doc->Save(ArtifactsDir + u"OoxmlSaveOptions.Iso29500Strict.docx", saveOptions);

        // The markup language of our shape has changed according to the compliance type
        doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.Iso29500Strict.docx");

        ASSERT_EQ(ShapeMarkupLanguage::Dml, (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_MarkupLanguage());
        //ExEnd
    }

    void RestartingDocumentList(bool doRestartListAtEachSection)
    {
        //ExStart
        //ExFor:List.IsRestartAtEachSection
        //ExSummary:Shows how to specify that the list has to be restarted at each section.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        doc->get_Lists()->Add(ListTemplate::NumberDefault);

        SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->idx_get(0);

        // Set true to specify that the list has to be restarted at each section
        list->set_IsRestartAtEachSection(doRestartListAtEachSection);

        // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376
        auto options = MakeObject<OoxmlSaveOptions>();
        options->set_Compliance(OoxmlCompliance::Iso29500_2008_Transitional);

        builder->get_ListFormat()->set_List(list);

        builder->Writeln(u"List item 1");
        builder->Writeln(u"List item 2");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Writeln(u"List item 3");
        builder->Writeln(u"List item 4");

        doc->Save(ArtifactsDir + u"OoxmlSaveOptions.RestartingDocumentList.docx", options);
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.RestartingDocumentList.docx");

        ASPOSE_ASSERT_EQ(doRestartListAtEachSection, doc->get_Lists()->idx_get(0)->get_IsRestartAtEachSection());
    }

    void UpdatingLastSavedTimeDocument()
    {
        //ExStart
        //ExFor:SaveOptions.UpdateLastSavedTimeProperty
        //ExSummary:Shows how to update a document time property when you want to save it.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Get last saved time
        System::DateTime documentTimeBeforeSave = doc->get_BuiltInDocumentProperties()->get_LastSavedTime();

        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_UpdateLastSavedTimeProperty(true);

        doc->Save(ArtifactsDir + u"OoxmlSaveOptions.UpdatingLastSavedTimeDocument.docx", saveOptions);
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        System::DateTime documentTimeAfterSave = doc->get_BuiltInDocumentProperties()->get_LastSavedTime();

        ASSERT_TRUE(documentTimeBeforeSave < documentTimeAfterSave);
    }

    void KeepLegacyControlChars(bool doKeepLegacyControlChars)
    {
        //ExStart
        //ExFor:OoxmlSaveOptions.KeepLegacyControlChars
        //ExFor:OoxmlSaveOptions.#ctor(SaveFormat)
        //ExSummary:Shows how to support legacy control characters when converting to .docx.
        auto doc = MakeObject<Document>(MyDir + u"Legacy control character.doc");

        // Note that only one legacy character (ShortDateTime) is supported which declared in the "DOC" format
        auto so = MakeObject<OoxmlSaveOptions>(SaveFormat::Docx);
        so->set_KeepLegacyControlChars(doKeepLegacyControlChars);

        doc->Save(ArtifactsDir + u"OoxmlSaveOptions.KeepLegacyControlChars.docx", so);

        // Open the saved document and verify results
        doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.KeepLegacyControlChars.docx");

        if (doKeepLegacyControlChars)
        {
            ASSERT_EQ(u"\u0013date \\@ \"MM/dd/yyyy\"\u0014\u0015\f", doc->get_FirstSection()->get_Body()->GetText());
        }
        else
        {
            ASSERT_EQ(u"\u001e\f", doc->get_FirstSection()->get_Body()->GetText());
        }
        //ExEnd
    }

    void DocumentCompression()
    {
        //ExStart
        //ExFor:OoxmlSaveOptions.CompressionLevel
        //ExFor:CompressionLevel
        //ExSummary:Shows how to specify the compression level used to save the OOXML document.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto saveOptions = MakeObject<OoxmlSaveOptions>(SaveFormat::Docx);
        // DOCX and DOTX files are internally a ZIP-archive, this property controls
        // the compression level of the archive
        // Note, that FlatOpc file is not a ZIP-archive, therefore, this property does
        // not affect the FlatOpc files
        // Aspose.Words uses CompressionLevel.Normal by default, but MS Word uses
        // CompressionLevel.SuperFast by default
        saveOptions->set_CompressionLevel(CompressionLevel::SuperFast);

        doc->Save(ArtifactsDir + u"OoxmlSaveOptions.out.docx", saveOptions);
        //ExEnd
    }

    void CheckFileSignatures()
    {
        ArrayPtr<CompressionLevel> compressionLevels =
            MakeArray<CompressionLevel>({CompressionLevel::Maximum, CompressionLevel::Normal, CompressionLevel::Fast, CompressionLevel::SuperFast});

        ArrayPtr<String> fileSignatures = MakeArray<String>(
            {u"50 4B 03 04 14 00 02 00 08 00 ", u"50 4B 03 04 14 00 00 00 08 00 ", u"50 4B 03 04 14 00 04 00 08 00 ", u"50 4B 03 04 14 00 06 00 08 00 "});

        auto doc = MakeObject<Document>();
        auto saveOptions = MakeObject<OoxmlSaveOptions>(SaveFormat::Docx);

        int64_t prevFileSize = 0;
        for (int i = 0; i < fileSignatures->get_Length(); i++)
        {
            saveOptions->set_CompressionLevel(compressionLevels[i]);
            doc->Save(ArtifactsDir + u"OoxmlSaveOptions.CheckFileSignatures.docx", saveOptions);

            {
                auto stream = MakeObject<System::IO::MemoryStream>();
                {
                    SharedPtr<System::IO::FileStream> outputFileStream =
                        System::IO::File::Open(ArtifactsDir + u"OoxmlSaveOptions.CheckFileSignatures.docx", System::IO::FileMode::Open);
                    int64_t fileSize = outputFileStream->get_Length();
                    ASSERT_TRUE(prevFileSize < fileSize);

                    TestUtil::CopyStream(outputFileStream, stream);
                    ASSERT_EQ(fileSignatures[i], TestUtil::DumpArray(stream->ToArray(), 0, 10));

                    prevFileSize = fileSize;
                }
            }
        }
    }
};

} // namespace ApiExamples
