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
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeMarkupLanguage.h>
#include <Aspose.Words.Cpp/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Lists/List.h>
#include <Aspose.Words.Cpp/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/CompressionLevel.h>
#include <Aspose.Words.Cpp/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Settings/MsWordVersion.h>
#include <system/array.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/diagnostics/stopwatch.h>
#include <system/enum.h>
#include <system/exceptions.h>
#include <system/io/file.h>
#include <system/io/file_info.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Loading;
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
        //ExSummary:Shows how to create a password encrypted Office Open XML document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_Password(u"MyPassword");

        doc->Save(ArtifactsDir + u"OoxmlSaveOptions.Password.docx", saveOptions);

        // We will not be able to open this document with Microsoft Word or
        // Aspose.Words without providing the correct password.
        ASSERT_THROW(doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.Password.docx"), IncorrectPasswordException);

        // Open the encrypted document by passing the correct password in a LoadOptions object.
        doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.Password.docx", MakeObject<LoadOptions>(u"MyPassword"));

        ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
        //ExEnd
    }

    void Iso29500Strict()
    {
        //ExStart
        //ExFor:CompatibilityOptions
        //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
        //ExFor:OoxmlSaveOptions
        //ExFor:OoxmlSaveOptions.#ctor()
        //ExFor:OoxmlSaveOptions.SaveFormat
        //ExFor:OoxmlCompliance
        //ExFor:OoxmlSaveOptions.Compliance
        //ExFor:ShapeMarkupLanguage
        //ExSummary:Shows how to set an OOXML compliance specification for a saved document to adhere to.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // If we configure compatibility options to comply with Microsoft Word 2003,
        // inserting an image will define its shape using VML.
        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2003);
        builder->InsertImage(ImageDir + u"Transparent background logo.png");

        ASSERT_EQ(ShapeMarkupLanguage::Vml, (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_MarkupLanguage());

        // The "ISO/IEC 29500:2008" OOXML standard does not support VML shapes.
        // If we set the "Compliance" property of the SaveOptions object to "OoxmlCompliance.Iso29500_2008_Strict",
        // any document we save while passing this object will have to follow that standard.
        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_Compliance(OoxmlCompliance::Iso29500_2008_Strict);
        saveOptions->set_SaveFormat(SaveFormat::Docx);

        doc->Save(ArtifactsDir + u"OoxmlSaveOptions.Iso29500Strict.docx", saveOptions);

        // Our saved document defines the shape using DML to adhere to the "ISO/IEC 29500:2008" OOXML standard.
        doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.Iso29500Strict.docx");

        ASSERT_EQ(ShapeMarkupLanguage::Dml, (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_MarkupLanguage());
        //ExEnd
    }

    void RestartingDocumentList(bool restartListAtEachSection)
    {
        //ExStart
        //ExFor:List.IsRestartAtEachSection
        //ExFor:OoxmlCompliance
        //ExFor:OoxmlSaveOptions.Compliance
        //ExSummary:Shows how to configure a list to restart numbering at each section.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        doc->get_Lists()->Add(ListTemplate::NumberDefault);

        SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->idx_get(0);
        list->set_IsRestartAtEachSection(restartListAtEachSection);

        // The "IsRestartAtEachSection" property will only be applicable when
        // the document's OOXML compliance level is to a standard that is newer than "OoxmlComplianceCore.Ecma376".
        auto options = MakeObject<OoxmlSaveOptions>();
        options->set_Compliance(OoxmlCompliance::Iso29500_2008_Transitional);

        builder->get_ListFormat()->set_List(list);

        builder->Writeln(u"List item 1");
        builder->Writeln(u"List item 2");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Writeln(u"List item 3");
        builder->Writeln(u"List item 4");

        doc->Save(ArtifactsDir + u"OoxmlSaveOptions.RestartingDocumentList.docx", options);

        doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.RestartingDocumentList.docx");

        ASPOSE_ASSERT_EQ(restartListAtEachSection, doc->get_Lists()->idx_get(0)->get_IsRestartAtEachSection());
        //ExEnd
    }

    void LastSavedTime(bool updateLastSavedTimeProperty)
    {
        //ExStart
        //ExFor:SaveOptions.UpdateLastSavedTimeProperty
        //ExSummary:Shows how to determine whether to preserve the document's "Last saved time" property when saving.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        ASSERT_EQ(System::DateTime(2021, 5, 11, 6, 32, 0), doc->get_BuiltInDocumentProperties()->get_LastSavedTime());

        // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
        // and then pass it to the document's saving method to modify how we save the document.
        // Set the "UpdateLastSavedTimeProperty" property to "true" to
        // set the output document's "Last saved time" built-in property to the current date/time.
        // Set the "UpdateLastSavedTimeProperty" property to "false" to
        // preserve the original value of the input document's "Last saved time" built-in property.
        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_UpdateLastSavedTimeProperty(updateLastSavedTimeProperty);

        doc->Save(ArtifactsDir + u"OoxmlSaveOptions.LastSavedTime.docx", saveOptions);

        doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.LastSavedTime.docx");
        System::DateTime lastSavedTimeNew = doc->get_BuiltInDocumentProperties()->get_LastSavedTime();

        if (updateLastSavedTimeProperty)
        {
            ASSERT_GT(System::DateTime::get_Now(), lastSavedTimeNew.AddDays(-1));
            ASSERT_LT(System::DateTime::get_Now(), lastSavedTimeNew.AddDays(1));
        }
        else
        {
            ASSERT_EQ(System::DateTime(2021, 5, 11, 6, 32, 0), lastSavedTimeNew);
        }
        //ExEnd
    }

    void KeepLegacyControlChars(bool keepLegacyControlChars)
    {
        //ExStart
        //ExFor:OoxmlSaveOptions.KeepLegacyControlChars
        //ExFor:OoxmlSaveOptions.#ctor(SaveFormat)
        //ExSummary:Shows how to support legacy control characters when converting to .docx.
        auto doc = MakeObject<Document>(MyDir + u"Legacy control character.doc");

        // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
        // and then pass it to the document's saving method to modify how we save the document.
        // Set the "KeepLegacyControlChars" property to "true" to preserve
        // the "ShortDateTime" legacy character while saving.
        // Set the "KeepLegacyControlChars" property to "false" to remove
        // the "ShortDateTime" legacy character from the output document.
        auto so = MakeObject<OoxmlSaveOptions>(SaveFormat::Docx);
        so->set_KeepLegacyControlChars(keepLegacyControlChars);

        doc->Save(ArtifactsDir + u"OoxmlSaveOptions.KeepLegacyControlChars.docx", so);

        doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.KeepLegacyControlChars.docx");

        ASSERT_EQ(keepLegacyControlChars ? String(u"\u0013date \\@ \"MM/dd/yyyy\"\u0014\u0015\f") : String(u"\u001e\f"),
                  doc->get_FirstSection()->get_Body()->GetText());
        //ExEnd
    }

    void DocumentCompression(CompressionLevel compressionLevel)
    {
        //ExStart
        //ExFor:OoxmlSaveOptions.CompressionLevel
        //ExFor:CompressionLevel
        //ExSummary:Shows how to specify the compression level to use while saving an OOXML document.
        auto doc = MakeObject<Document>(MyDir + u"Big document.docx");

        // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
        // and then pass it to the document's saving method to modify how we save the document.
        // Set the "CompressionLevel" property to "CompressionLevel.Maximum" to apply the strongest and slowest compression.
        // Set the "CompressionLevel" property to "CompressionLevel.Normal" to apply
        // the default compression that Aspose.Words uses while saving OOXML documents.
        // Set the "CompressionLevel" property to "CompressionLevel.Fast" to apply a faster and weaker compression.
        // Set the "CompressionLevel" property to "CompressionLevel.SuperFast" to apply
        // the default compression that Microsoft Word uses.
        auto saveOptions = MakeObject<OoxmlSaveOptions>(SaveFormat::Docx);
        saveOptions->set_CompressionLevel(compressionLevel);

        SharedPtr<System::Diagnostics::Stopwatch> st = System::Diagnostics::Stopwatch::StartNew();
        doc->Save(ArtifactsDir + u"OoxmlSaveOptions.DocumentCompression.docx", saveOptions);
        st->Stop();

        auto fileInfo = MakeObject<System::IO::FileInfo>(ArtifactsDir + u"OoxmlSaveOptions.DocumentCompression.docx");

        std::cout << String::Format(u"Saving operation done using the \"{0}\" compression level:", compressionLevel) << std::endl;
        std::cout << "\tDuration:\t" << st->get_ElapsedMilliseconds() << " ms" << std::endl;
        std::cout << "\tFile Size:\t" << fileInfo->get_Length() << " bytes" << std::endl;
        //ExEnd

        switch (compressionLevel)
        {
        case CompressionLevel::Maximum:
            ASSERT_GE(1266000, fileInfo->get_Length());
            break;

        case CompressionLevel::Normal:
            ASSERT_LT(1266900, fileInfo->get_Length());
            break;

        case CompressionLevel::Fast:
            ASSERT_LT(1269000, fileInfo->get_Length());
            break;

        case CompressionLevel::SuperFast:
            ASSERT_LT(1271000, fileInfo->get_Length());
            break;
        }
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
