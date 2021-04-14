#pragma once

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Saving/MarkdownSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/TableContentAlignment.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <system/details/dispose_guard.h>
#include <system/io/memory_stream.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options {

class WorkingWithMarkdownSaveOptions : public DocsExamplesBase
{
public:
    void ExportIntoMarkdownWithTableContentAlignment()
    {
        //ExStart:ExportIntoMarkdownWithTableContentAlignment
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertCell();
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);
        builder->Write(u"Cell1");
        builder->InsertCell();
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
        builder->Write(u"Cell2");

        // Makes all paragraphs inside the table to be aligned.
        auto saveOptions = MakeObject<MarkdownSaveOptions>();
        saveOptions->set_TableContentAlignment(TableContentAlignment::Left);
        doc->Save(ArtifactsDir + u"WorkingWithMarkdownSaveOptions.LeftTableContentAlignment.md", saveOptions);

        saveOptions->set_TableContentAlignment(TableContentAlignment::Right);
        doc->Save(ArtifactsDir + u"WorkingWithMarkdownSaveOptions.RightTableContentAlignment.md", saveOptions);

        saveOptions->set_TableContentAlignment(TableContentAlignment::Center);
        doc->Save(ArtifactsDir + u"WorkingWithMarkdownSaveOptions.CenterTableContentAlignment.md", saveOptions);

        // The alignment in this case will be taken from the first paragraph in corresponding table column.
        saveOptions->set_TableContentAlignment(TableContentAlignment::Auto);
        doc->Save(ArtifactsDir + u"WorkingWithMarkdownSaveOptions.AutoTableContentAlignment.md", saveOptions);
        //ExEnd:ExportIntoMarkdownWithTableContentAlignment
    }

    void SetImagesFolder()
    {
        //ExStart:SetImagesFolder
        auto doc = MakeObject<Document>(MyDir + u"Image bullet points.docx");

        auto saveOptions = MakeObject<MarkdownSaveOptions>();
        saveOptions->set_ImagesFolder(ArtifactsDir + u"Images");

        {
            auto stream = MakeObject<System::IO::MemoryStream>();
            doc->Save(stream, saveOptions);
        }
        //ExEnd:SetImagesFolder
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
