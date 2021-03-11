#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

void ReplaceTextInTable()
{
    System::String inputDataDir = GetInputDataDir_FindAndReplace();
    System::String outputDataDir = GetOutputDataDir_FindAndReplace();
    //ExStart:ReplaceText
    auto doc = System::MakeObject<Document>(inputDataDir + u"Tables.docx");
    auto table = System::DynamicCast<Tables::Table>(doc->GetChild(NodeType::Table, 0, true));

    table->get_Range()->Replace(u"Carrots", u"Eggs", System::MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));
    table->get_LastRow()->get_LastCell()->get_Range()->Replace(u"50", u"20", System::MakeObject<FindReplaceOptions>(FindReplaceDirection::Forward));

    doc->Save(outputDataDir + u"FindAndReplace.ReplaceTextInTable.docx");
    //ExEnd:ReplaceText
}
