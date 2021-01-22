#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Markup/MarkupLevel.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtType.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <drawing/color.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;

void RichTextBoxContentControl()
{
    std::cout << "RichTextBoxContentControl example started." << std::endl;
    // ExStart:RichTextBoxContentControl
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<StructuredDocumentTag> sdtRichText = System::MakeObject<StructuredDocumentTag>(doc, SdtType::RichText, MarkupLevel::Block);

    System::SharedPtr<Paragraph> para = System::MakeObject<Paragraph>(doc);
    System::SharedPtr<Run> run = System::MakeObject<Run>(doc);
    run->set_Text(u"Hello World");
    run->get_Font()->set_Color(System::Drawing::Color::get_Green());
    para->get_Runs()->Add(run);
    sdtRichText->get_ChildNodes()->Add(para);
    doc->get_FirstSection()->get_Body()->AppendChild(sdtRichText);

    System::String outputPath = outputDataDir + u"RichTextBoxContentControl.docx";
    doc->Save(outputPath);
    // ExEnd:RichTextBoxContentControl
    std::cout << "Rich text box type content control created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "RichTextBoxContentControl example finished." << std::endl << std::endl;
}