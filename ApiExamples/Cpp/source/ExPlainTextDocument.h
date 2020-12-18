#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/PlainTextDocument.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Properties/CustomDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Properties/DocumentProperty.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <system/details/dispose_guard.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExPlainTextDocument : public ApiExampleBase
{
public:
    void Load()
    {
        //ExStart
        //ExFor:PlainTextDocument
        //ExFor:PlainTextDocument.#ctor(String)
        //ExFor:PlainTextDocument.Text
        //ExSummary:Shows how to load a document in its plaintext state.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");
        doc->Save(ArtifactsDir + u"PlainTextDocument.Load.docx");

        auto plaintext = MakeObject<PlainTextDocument>(ArtifactsDir + u"PlainTextDocument.Load.docx");

        ASSERT_EQ(u"Hello world!", plaintext->get_Text().Trim());
        //ExEnd
    }

    void LoadWithOptions()
    {
        //ExStart
        //ExFor:PlainTextDocument.#ctor(String, LoadOptions)
        //ExSummary:Shows how to load an encrypted document in its plaintext state.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");

        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_Password(u"MyPassword");

        doc->Save(ArtifactsDir + u"PlainTextDocument.LoadWithOptions.docx", saveOptions);

        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_Password(u"MyPassword");

        auto plaintext = MakeObject<PlainTextDocument>(ArtifactsDir + u"PlainTextDocument.LoadWithOptions.docx", loadOptions);

        ASSERT_EQ(u"Hello world!", plaintext->get_Text().Trim());
        //ExEnd
    }

    void LoadFromStream()
    {
        //ExStart
        //ExFor:PlainTextDocument.#ctor(Stream)
        //ExSummary:Shows how to load a document from a stream in its plaintext state.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");
        doc->Save(ArtifactsDir + u"PlainTextDocument.LoadFromStream.docx");

        {
            auto stream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"PlainTextDocument.LoadFromStream.docx", System::IO::FileMode::Open);
            auto plaintext = MakeObject<PlainTextDocument>(stream);

            ASSERT_EQ(u"Hello world!", plaintext->get_Text().Trim());
        }
        //ExEnd
    }

    void LoadFromStreamWithOptions()
    {
        //ExStart
        //ExFor:PlainTextDocument.#ctor(Stream, LoadOptions)
        //ExSummary:Shows how to load an encrypted document from a stream in its plaintext state.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");

        auto saveOptions = MakeObject<OoxmlSaveOptions>();
        saveOptions->set_Password(u"MyPassword");

        doc->Save(ArtifactsDir + u"PlainTextDocument.LoadFromStreamWithOptions.docx", saveOptions);

        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_Password(u"MyPassword");

        {
            auto stream =
                MakeObject<System::IO::FileStream>(ArtifactsDir + u"PlainTextDocument.LoadFromStreamWithOptions.docx", System::IO::FileMode::Open);
            auto plaintext = MakeObject<PlainTextDocument>(stream, loadOptions);

            ASSERT_EQ(u"Hello world!", plaintext->get_Text().Trim());
        }
        //ExEnd
    }

    void BuiltInProperties()
    {
        //ExStart
        //ExFor:PlainTextDocument.BuiltInDocumentProperties
        //ExSummary:Shows how to load a plaintext version of a document, and also access its built in properties.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");
        doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");

        doc->Save(ArtifactsDir + u"PlainTextDocument.BuiltInProperties.docx");

        auto plaintext = MakeObject<PlainTextDocument>(ArtifactsDir + u"PlainTextDocument.BuiltInProperties.docx");

        ASSERT_EQ(u"Hello world!", plaintext->get_Text().Trim());
        ASSERT_EQ(u"John Doe", plaintext->get_BuiltInDocumentProperties()->get_Author());
        //ExEnd
    }

    void CustomDocumentProperties_()
    {
        //ExStart
        //ExFor:PlainTextDocument.CustomDocumentProperties
        //ExSummary:Shows how to load a plaintext version of a document, and also access its custom properties.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");
        doc->get_CustomDocumentProperties()->Add(u"Location of writing", String(u"123 Main St, London, UK"));

        doc->Save(ArtifactsDir + u"PlainTextDocument.CustomDocumentProperties.docx");

        auto plaintext = MakeObject<PlainTextDocument>(ArtifactsDir + u"PlainTextDocument.CustomDocumentProperties.docx");

        ASSERT_EQ(u"Hello world!", plaintext->get_Text().Trim());
        ASPOSE_ASSERT_EQ(u"123 Main St, London, UK", plaintext->get_CustomDocumentProperties()->idx_get(u"Location of writing")->get_Value());
        //ExEnd
    }
};

} // namespace ApiExamples
