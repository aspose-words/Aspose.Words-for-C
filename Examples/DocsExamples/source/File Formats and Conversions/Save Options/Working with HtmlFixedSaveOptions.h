#pragma once

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Saving/HtmlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options {

class WorkingWithHtmlFixedSaveOptions : public DocsExamplesBase
{
public:
    void UseFontFromTargetMachine()
    {
        //ExStart:UseFontFromTargetMachine
        auto doc = MakeObject<Document>(MyDir + u"Bullet points with alternative font.docx");

        auto saveOptions = MakeObject<HtmlFixedSaveOptions>();
        saveOptions->set_UseTargetMachineFonts(true);

        doc->Save(ArtifactsDir + u"WorkingWithHtmlFixedSaveOptions.UseFontFromTargetMachine.html", saveOptions);
        //ExEnd:UseFontFromTargetMachine
    }

    void WriteAllCssRulesInSingleFile()
    {
        //ExStart:WriteAllCssRulesInSingleFile
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Setting this property to true restores the old behavior (separate files) for compatibility with legacy code.
        // All CSS rules are written into single file "styles.css.
        auto saveOptions = MakeObject<HtmlFixedSaveOptions>();
        saveOptions->set_SaveFontFaceCssSeparately(false);

        doc->Save(ArtifactsDir + u"WorkingWithHtmlFixedSaveOptions.WriteAllCssRulesInSingleFile.html", saveOptions);
        //ExEnd:WriteAllCssRulesInSingleFile
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
