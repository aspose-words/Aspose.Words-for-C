#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/TxtListIndentation.h>
#include <Aspose.Words.Cpp/Model/Saving/TxtSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options {

class WorkingWithTxtSaveOptions : public DocsExamplesBase
{
public:
    void AddBidiMarks()
    {
        //ExStart:AddBidiMarks
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");
        builder->get_ParagraphFormat()->set_Bidi(true);
        builder->Writeln(u"שלום עולם!");
        builder->Writeln(u"مرحبا بالعالم!");

        auto saveOptions = MakeObject<TxtSaveOptions>();
        saveOptions->set_AddBidiMarks(true);

        doc->Save(ArtifactsDir + u"WorkingWithTxtSaveOptions.AddBidiMarks.txt", saveOptions);
        //ExEnd:AddBidiMarks
    }

    void UseTabCharacterPerLevelForListIndentation()
    {
        //ExStart:UseTabCharacterPerLevelForListIndentation
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a list with three levels of indentation.
        builder->get_ListFormat()->ApplyNumberDefault();
        builder->Writeln(u"Item 1");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->ListIndent();
        builder->Write(u"Item 3");

        auto saveOptions = MakeObject<TxtSaveOptions>();
        saveOptions->get_ListIndentation()->set_Count(1);
        saveOptions->get_ListIndentation()->set_Character(u'\t');

        doc->Save(ArtifactsDir + u"WorkingWithTxtSaveOptions.UseTabCharacterPerLevelForListIndentation.txt", saveOptions);
        //ExEnd:UseTabCharacterPerLevelForListIndentation
    }

    void UseSpaceCharacterPerLevelForListIndentation()
    {
        //ExStart:UseSpaceCharacterPerLevelForListIndentation
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a list with three levels of indentation.
        builder->get_ListFormat()->ApplyNumberDefault();
        builder->Writeln(u"Item 1");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->ListIndent();
        builder->Write(u"Item 3");

        auto saveOptions = MakeObject<TxtSaveOptions>();
        saveOptions->get_ListIndentation()->set_Count(3);
        saveOptions->get_ListIndentation()->set_Character(u' ');

        doc->Save(ArtifactsDir + u"WorkingWithTxtSaveOptions.UseSpaceCharacterPerLevelForListIndentation.txt", saveOptions);
        //ExEnd:UseSpaceCharacterPerLevelForListIndentation
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
