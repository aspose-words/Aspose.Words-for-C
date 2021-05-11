#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Saving/WordML2003SaveOptions.h>
#include <system/io/file.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExWordML2003SaveOptions : public ApiExampleBase
{
public:
    void PrettyFormat(bool prettyFormat)
    {
        //ExStart
        //ExFor:WordML2003SaveOptions
        //ExFor:WordML2003SaveOptions.SaveFormat
        //ExSummary:Shows how to manage output document's raw content.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // Create a "WordML2003SaveOptions" object to pass to the document's "Save" method
        // to modify how we save the document to the WordML save format.
        auto options = MakeObject<WordML2003SaveOptions>();

        ASSERT_EQ(SaveFormat::WordML, options->get_SaveFormat());

        // Set the "PrettyFormat" property to "true" to apply tab character indentation and
        // newlines to make the output document's raw content easier to read.
        // Set the "PrettyFormat" property to "false" to save the document's raw content in one continuous body of the text.
        options->set_PrettyFormat(prettyFormat);

        doc->Save(ArtifactsDir + u"WordML2003SaveOptions.PrettyFormat.xml", options);

        String fileContents = System::IO::File::ReadAllText(ArtifactsDir + u"WordML2003SaveOptions.PrettyFormat.xml");

        if (prettyFormat)
        {
            ASSERT_TRUE(fileContents.Contains(String(u"<o:DocumentProperties>\r\n\t\t") + u"<o:Revision>1</o:Revision>\r\n\t\t" +
                                              u"<o:TotalTime>0</o:TotalTime>\r\n\t\t" + u"<o:Pages>1</o:Pages>\r\n\t\t" + u"<o:Words>0</o:Words>\r\n\t\t" +
                                              u"<o:Characters>0</o:Characters>\r\n\t\t" + u"<o:Lines>1</o:Lines>\r\n\t\t" +
                                              u"<o:Paragraphs>1</o:Paragraphs>\r\n\t\t" + u"<o:CharactersWithSpaces>0</o:CharactersWithSpaces>\r\n\t\t" +
                                              u"<o:Version>11.5606</o:Version>\r\n\t" + u"</o:DocumentProperties>"));
        }
        else
        {
            ASSERT_TRUE(fileContents.Contains(String(u"<o:DocumentProperties><o:Revision>1</o:Revision><o:TotalTime>0</o:TotalTime><o:Pages>1</o:Pages>") +
                                              u"<o:Words>0</o:Words><o:Characters>0</o:Characters><o:Lines>1</o:Lines><o:Paragraphs>1</o:Paragraphs>" +
                                              u"<o:CharactersWithSpaces>0</o:CharactersWithSpaces><o:Version>11.5606</o:Version></o:DocumentProperties>"));
        }
        //ExEnd
    }

    void MemoryOptimization(bool memoryOptimization)
    {
        //ExStart
        //ExFor:WordML2003SaveOptions
        //ExSummary:Shows how to manage memory optimization.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // Create a "WordML2003SaveOptions" object to pass to the document's "Save" method
        // to modify how we save the document to the WordML save format.
        auto options = MakeObject<WordML2003SaveOptions>();

        // Set the "MemoryOptimization" flag to "true" to decrease memory consumption
        // during the document's saving operation at the cost of a longer saving time.
        // Set the "MemoryOptimization" flag to "false" to save the document normally.
        options->set_MemoryOptimization(memoryOptimization);

        doc->Save(ArtifactsDir + u"WordML2003SaveOptions.MemoryOptimization.xml", options);
        //ExEnd
    }
};

} // namespace ApiExamples
