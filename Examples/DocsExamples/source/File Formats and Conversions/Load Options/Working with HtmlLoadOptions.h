#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Loading/HtmlControlType.h>
#include <Aspose.Words.Cpp/Loading/HtmlLoadOptions.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/array.h>
#include <system/io/memory_stream.h>
#include <system/text/encoding.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Load_Options {

class WorkingWithHtmlLoadOptions : public DocsExamplesBase
{
public:
    void PreferredControlType()
    {
        //ExStart:LoadHtmlElementsWithPreferredControlType
        const String html = u"\r\n                <html>\r\n                    <select name='ComboBox' size='1'>\r\n                        <option "
                            u"value='val1'>item1</option>\r\n                        <option value='val2'></option>                        \r\n                "
                            u"    </select>\r\n                </html>\r\n            ";

        auto loadOptions = MakeObject<HtmlLoadOptions>();
        loadOptions->set_PreferredControlType(HtmlControlType::StructuredDocumentTag);

        auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(html)), loadOptions);

        doc->Save(ArtifactsDir + u"WorkingWithHtmlLoadOptions.PreferredControlType.docx", SaveFormat::Docx);
        //ExEnd:LoadHtmlElementsWithPreferredControlType
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Load_Options
