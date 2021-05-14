#include "Split into html pages.h"

#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/ImportFormatMode.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Saving/ExportHeadersFootersMode.h>
#include <Aspose.Words.Cpp/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <system/char.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/text/string_builder.h>

#include "Programming with Documents/Split Documents/Split into html pages.h"

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
using namespace Aspose::Words::Saving;
namespace DocsExamples { namespace Programming_with_Documents { namespace Split_Documents {

void WordToHtmlConverter::HandleTocMergeField::FieldMerging(System::SharedPtr<FieldMergingArgs> e)
{
    if (mBuilder == nullptr)
    {
        mBuilder = System::MakeObject<DocumentBuilder>(e->get_Document());
    }

    // Our custom data source returns topic objects.
    auto topic = System::StaticCast<Topic>(e->get_FieldValue());

    mBuilder->MoveToMergeField(e->get_FieldName());
    mBuilder->InsertHyperlink(topic->get_Title(), topic->get_FileName(), false);

    // Signal to the mail merge engine that it does not need to insert text into the field.
    e->set_Text(u"");
}

void WordToHtmlConverter::HandleTocMergeField::ImageFieldMerging(System::SharedPtr<ImageFieldMergingArgs> args)
{
    // Do nothing.
}

void WordToHtmlConverter::Execute(System::String srcFileName, System::String tocTemplate, System::String dstDir)
{
    mDoc = System::MakeObject<Document>(srcFileName);
    mTocTemplate = tocTemplate;
    mDstDir = dstDir;

    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Paragraph>>> topicStartParas = SelectTopicStarts();
    InsertSectionBreaks(topicStartParas);
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Topic>>> topics = SaveHtmlTopics();
    SaveTableOfContents(topics);
}

System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Paragraph>>> WordToHtmlConverter::SelectTopicStarts()
{
    System::SharedPtr<NodeCollection> paras = mDoc->GetChildNodes(NodeType::Paragraph, true);
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Paragraph>>> topicStartParas =
        System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Paragraph>>>();

    for (auto para : System::IterateOver<Paragraph>(paras))
    {
        StyleIdentifier style = para->get_ParagraphFormat()->get_StyleIdentifier();
        if (style == StyleIdentifier::Heading1)
        {
            topicStartParas->Add(para);
        }
    }

    return topicStartParas;
}

void WordToHtmlConverter::InsertSectionBreaks(System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Paragraph>>> topicStartParas)
{
    auto builder = System::MakeObject<DocumentBuilder>(mDoc);
    for (auto para : topicStartParas)
    {
        System::SharedPtr<Section> section = para->get_ParentSection();

        // Insert section break if the paragraph is not at the beginning of a section already.
        if (para != section->get_Body()->get_FirstParagraph())
        {
            builder->MoveTo(para->get_FirstChild());
            builder->InsertBreak(BreakType::SectionBreakNewPage);

            // This is the paragraph that was inserted at the end of the now old section.
            // We don't really need the extra paragraph, we just needed the section.
            section->get_Body()->get_LastParagraph()->Remove();
        }
    }
}

System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Topic>>> WordToHtmlConverter::SaveHtmlTopics()
{
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Topic>>> topics =
        System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Topic>>>();
    for (int32_t sectionIdx = 0; sectionIdx < mDoc->get_Sections()->get_Count(); sectionIdx++)
    {
        System::SharedPtr<Section> section = mDoc->get_Sections()->idx_get(sectionIdx);

        System::String paraText = section->get_Body()->get_FirstParagraph()->GetText();

        // Use the text of the heading paragraph to generate the HTML file name.
        System::String fileName = MakeTopicFileName(paraText);
        if (fileName == u"")
        {
            fileName = System::String(u"UNTITLED SECTION ") + sectionIdx;
        }

        fileName = System::IO::Path::Combine(mDstDir, fileName + u".html");

        // Use the text of the heading paragraph to generate the title for the TOC.
        System::String title = MakeTopicTitle(paraText);
        if (title == u"")
        {
            title = System::String(u"UNTITLED SECTION ") + sectionIdx;
        }

        auto topic = System::MakeObject<Topic>(title, fileName);
        topics->Add(topic);

        SaveHtmlTopic(section, topic);
    }

    return topics;
}

System::String WordToHtmlConverter::MakeTopicFileName(System::String paraText)
{
    auto b = System::MakeObject<System::Text::StringBuilder>();
    for (char16_t c : paraText)
    {
        if (System::Char::IsLetterOrDigit(c))
        {
            b->Append(c);
        }
        else if (c == u' ')
        {
            b->Append(u'_');
        }
    }

    return b->ToString();
}

System::String WordToHtmlConverter::MakeTopicTitle(System::String paraText)
{
    return paraText.Substring(0, paraText.get_Length() - 1);
}

void WordToHtmlConverter::SaveHtmlTopic(System::SharedPtr<Section> section, System::SharedPtr<Topic> topic)
{
    auto dummyDoc = System::MakeObject<Document>();
    dummyDoc->RemoveAllChildren();
    dummyDoc->AppendChild(dummyDoc->ImportNode(section, true, ImportFormatMode::KeepSourceFormatting));

    dummyDoc->get_BuiltInDocumentProperties()->set_Title(topic->get_Title());

    auto saveOptions = System::MakeObject<HtmlSaveOptions>();
    saveOptions->set_PrettyFormat(true);
    saveOptions->set_AllowNegativeIndent(true);
    saveOptions->set_ExportHeadersFootersMode(ExportHeadersFootersMode::None);

    dummyDoc->Save(topic->get_FileName(), saveOptions);
}

void WordToHtmlConverter::SaveTableOfContents(System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Topic>>> topics)
{
    auto tocDoc = System::MakeObject<Document>(mTocTemplate);

    // We use a custom mail merge event handler defined below,
    // and a custom mail merge data source based on collecting the topics we created.
    tocDoc->get_MailMerge()->set_FieldMergingCallback(System::MakeObject<WordToHtmlConverter::HandleTocMergeField>());
    tocDoc->get_MailMerge()->ExecuteWithRegions(System::MakeObject<TocMailMergeDataSource>(topics));

    tocDoc->Save(System::IO::Path::Combine(mDstDir, u"contents.html"));
}

namespace gtest_test {

class SplitIntoHtmlPages : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Split_Documents::SplitIntoHtmlPages> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Split_Documents::SplitIntoHtmlPages>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Split_Documents::SplitIntoHtmlPages> SplitIntoHtmlPages::s_instance;

TEST_F(SplitIntoHtmlPages, HtmlPages)
{
    s_instance->HtmlPages();
}

} // namespace gtest_test

}}} // namespace DocsExamples::Programming_with_Documents::Split_Documents
