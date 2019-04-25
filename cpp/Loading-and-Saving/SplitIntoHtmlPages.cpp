#include "stdafx.h"
#include "examples.h"

#include <Model/Document/BreakType.h>
#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/MailMerge/MailMerge.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Properties/BuiltInDocumentProperties.h>
#include <Model/Saving/ExportHeadersFootersMode.h>
#include <Model/Saving/HtmlSaveOptions.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/SectionCollection.h>
#include <Model/Styles/StyleIdentifier.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/MailMerge/FieldMergingArgs.h>
#include <Model/MailMerge/IFieldMergingCallback.h>
#include <Model/MailMerge/ImageFieldMergingArgs.h>
#include <Model/MailMerge/IMailMergeDataSource.h>

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
using namespace Aspose::Words::Saving;

namespace
{
    class Topic;

    typedef System::SharedPtr<Paragraph> TParagraphPtr;
    typedef System::SharedPtr<Topic> TTopicPtr;

    class Topic : public System::Object
    {
        typedef Topic ThisType;
        typedef System::Object BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        Topic(System::String const &title, System::String const &filename);
        System::String GetTitle();
        System::String GetFilename();

    private:
        System::String mTitle;
        System::String mFilename;
    };

    RTTI_INFO_IMPL_HASH(3130410972u, Topic, ThisTypeBaseTypesInfo);

    Topic::Topic(System::String const &title, System::String const &filename)
    {
        mTitle = title;
        mFilename = filename;
    }

    System::String Topic::GetTitle()
    {
        return mTitle;
    }

    System::String Topic::GetFilename()
    {
        return mFilename;
    }

    class TocMailMergeDataSource : public IMailMergeDataSource
    {
        typedef TocMailMergeDataSource ThisType;
        typedef IMailMergeDataSource BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        TocMailMergeDataSource(std::vector<TTopicPtr> const & topics);
        System::String get_TableName();
        bool MoveNext();
        bool GetValue(System::String fieldName, System::SharedPtr<System::Object>& fieldValue);
        System::SharedPtr<IMailMergeDataSource> GetChildDataSource(System::String tableName);

    private:
        std::vector<TTopicPtr> mTopics;
        int32_t mIndex;
    };

    RTTI_INFO_IMPL_HASH(270249105u, TocMailMergeDataSource, ThisTypeBaseTypesInfo);

    TocMailMergeDataSource::TocMailMergeDataSource(std::vector<TTopicPtr> const &topics) : mIndex(-1), mTopics(topics)
    {
    }

    System::String TocMailMergeDataSource::get_TableName()
    {
        return u"TOC";
    }

    bool TocMailMergeDataSource::MoveNext()
    {
        if (mIndex < mTopics.size() - 1)
        {
            mIndex++;
            return true;
        }
        else
        {
            // Reached EOF, return false.
            return false;
        }
    }

    bool TocMailMergeDataSource::GetValue(System::String fieldName, System::SharedPtr<System::Object>& fieldValue)
    {
        if (fieldName == u"TocEntry")
        {
            // The template document is supposed to have only one field called "TocEntry".
            fieldValue = mTopics[mIndex];
            return true;
        }
        else
        {
            fieldValue.reset();
            return false;
        }
    }

    System::SharedPtr<IMailMergeDataSource> TocMailMergeDataSource::GetChildDataSource(System::String tableName)
    {
        return nullptr;
    }

    class HandleTocMergeField : public IFieldMergingCallback
    {
        typedef HandleTocMergeField ThisType;
        typedef IFieldMergingCallback BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        void FieldMerging(System::SharedPtr<FieldMergingArgs> e);
        void ImageFieldMerging(System::SharedPtr<ImageFieldMergingArgs> args);

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        System::SharedPtr<DocumentBuilder> mBuilder;
    };

    RTTI_INFO_IMPL_HASH(2701695377u, HandleTocMergeField, ThisTypeBaseTypesInfo);

    void HandleTocMergeField::FieldMerging(System::SharedPtr<FieldMergingArgs> e)
    {
        if (mBuilder == nullptr)
        {
            mBuilder = System::MakeObject<DocumentBuilder>(e->get_Document());
        }

        // Our custom data source returns topic objects.
        TTopicPtr topic = System::StaticCast<Topic>(e->get_FieldValue());

        // We use the document builder to move to the current merge field and insert a hyperlink.
        mBuilder->MoveToMergeField(e->get_FieldName());
        mBuilder->InsertHyperlink(topic->GetTitle(), topic->GetFilename(), false);

        // Signal to the mail merge engine that it does not need to insert text into the field
        // As we've done it already.
        e->set_Text(u"");
    }

    void HandleTocMergeField::ImageFieldMerging(System::SharedPtr<ImageFieldMergingArgs> args) { }

    System::Object::shared_members_type HandleTocMergeField::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("HandleTocMergeField::mBuilder", this->mBuilder);
        return result;
    }

    class Worker : public System::Object
    {
        typedef Worker ThisType;
        typedef System::Object BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        void Execute(System::String const &srcFileName, System::String const &tocTemplate, System::String const &dstDir);

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        System::SharedPtr<Document> mDoc;
        System::String mTocTemplate;
        System::String mDstDir;

        std::vector<TParagraphPtr> SelectTopicStarts();
        void InsertSectionBreaks(std::vector<TParagraphPtr> const &topicStartParas);
        std::vector<TTopicPtr> SaveHtmlTopics();
        static System::String MakeTopicFileName(System::String paraText);
        static System::String MakeTopicTitle(System::String paraText);
        static void SaveHtmlTopic(System::SharedPtr<Section> section, TTopicPtr topic);
        void SaveTableOfContents(std::vector<TTopicPtr> const &topics);
    };

    RTTI_INFO_IMPL_HASH(3546846289u, Worker, ThisTypeBaseTypesInfo);

    void Worker::Execute(System::String const &srcFileName, System::String  const &tocTemplate, System::String  const &dstDir)
    {
        mDoc = System::MakeObject<Document>(srcFileName);
        mTocTemplate = tocTemplate;
        mDstDir = dstDir;

        std::vector<TParagraphPtr> topicStartParas = SelectTopicStarts();
        InsertSectionBreaks(topicStartParas);
        std::vector<TTopicPtr> topics = SaveHtmlTopics();
        SaveTableOfContents(topics);
    }

    std::vector<TParagraphPtr> Worker::SelectTopicStarts()
    {
        System::SharedPtr<NodeCollection> paras = mDoc->GetChildNodes(NodeType::Paragraph, true);
        std::vector<TParagraphPtr> topicStartParas;

        for (TParagraphPtr para : System::IterateOver<Paragraph>(paras))
        {
            StyleIdentifier style = para->get_ParagraphFormat()->get_StyleIdentifier();
            if (style == StyleIdentifier::Heading1)
            {
                topicStartParas.push_back(para);
            }
        }

        return topicStartParas;
    }

    void Worker::InsertSectionBreaks(std::vector<TParagraphPtr> const &topicStartParas)
    {
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(mDoc);
        for (TParagraphPtr para : topicStartParas)
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

    std::vector<TTopicPtr> Worker::SaveHtmlTopics()
    {
        std::vector<TTopicPtr> topics;
        for (int32_t sectionIdx = 0; sectionIdx < mDoc->get_Sections()->get_Count(); ++sectionIdx)
        {
            System::SharedPtr<Section> section = mDoc->get_Sections()->idx_get(sectionIdx);

            System::String paraText = section->get_Body()->get_FirstParagraph()->GetText();

            // The text of the heading paragaph is used to generate the HTML file name.
            System::String fileName = MakeTopicFileName(paraText);
            if (fileName == u"")
            {
                fileName = System::String(u"UNTITLED SECTION ") + sectionIdx;
            }

            fileName = System::IO::Path::Combine(mDstDir, fileName + u".html");

            // The text of the heading paragraph is also used to generate the title for the TOC.
            System::String title = MakeTopicTitle(paraText);
            if (title == u"")
            {
                title = System::String(u"UNTITLED SECTION ") + sectionIdx;
            }

            TTopicPtr topic = System::MakeObject<Topic>(title, fileName);
            topics.push_back(topic);

            SaveHtmlTopic(section, topic);
        }

        return topics;
    }

    System::String Worker::MakeTopicFileName(System::String paraText)
    {
        System::SharedPtr<System::Text::StringBuilder> b = System::MakeObject<System::Text::StringBuilder>();
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

    System::String Worker::MakeTopicTitle(System::String paraText)
    {
        return paraText.Substring(0, paraText.get_Length() - 1);
    }

    void Worker::SaveHtmlTopic(System::SharedPtr<Section> section, TTopicPtr topic)
    {
        System::SharedPtr<Document> dummyDoc = System::MakeObject<Document>();
        dummyDoc->RemoveAllChildren();
        dummyDoc->AppendChild(dummyDoc->ImportNode(section, true, ImportFormatMode::KeepSourceFormatting));

        dummyDoc->get_BuiltInDocumentProperties()->set_Title(topic->GetTitle());

        System::SharedPtr<HtmlSaveOptions> saveOptions = System::MakeObject<HtmlSaveOptions>();
        saveOptions->set_PrettyFormat(true);
        // This is to allow headings to appear to the left of main text.
        saveOptions->set_AllowNegativeIndent(true);
        saveOptions->set_ExportHeadersFootersMode(ExportHeadersFootersMode::None);

        dummyDoc->Save(topic->GetFilename(), saveOptions);
    }

    void Worker::SaveTableOfContents(std::vector<TTopicPtr> const &topics)
    {
        System::SharedPtr<Document> tocDoc = System::MakeObject<Document>(mTocTemplate);

        // We use a custom mail merge even handler defined below.
        tocDoc->get_MailMerge()->set_FieldMergingCallback(System::MakeObject<HandleTocMergeField>());
        // We use a custom mail merge data source based on the collection of the topics we created.
        tocDoc->get_MailMerge()->ExecuteWithRegions(System::MakeObject<TocMailMergeDataSource>(topics));

        tocDoc->Save(System::IO::Path::Combine(mDstDir, u"contents.html"));
    }

    System::Object::shared_members_type Worker::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("Worker::mDoc", this->mDoc);
        return result;
    }
}

void SplitIntoHtmlPages()
{
    std::cout << "SplitIntoHtmlPages example started." << std::endl;
    // You need to have a valid license for Aspose.Words.
    // The best way is to embed the license as a resource into the project
    // And specify only file name without path in the following call.
    // Aspose.Words.License license = new Aspose.Words.License();
    // License.SetLicense(@"Aspose.Words.lic");

    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();

    System::String srcFileName = dataDir + u"SOI 2007-2012-DeeM with footnote added.doc";
    System::String tocTemplate = dataDir + u"TocTemplate.doc";

    System::String outDir = System::IO::Path::Combine(dataDir, u"_out");
    System::IO::Directory::CreateDirectory_(outDir);

    // This class does the job.
    System::SharedPtr<Worker> w = System::MakeObject<Worker>();
    w->Execute(srcFileName, tocTemplate, outDir);

    std::cout << "Document split into HTML pages successfully." << std::endl << "Files saved at " << outDir.ToUtf8String() << std::endl;
    std::cout << "SplitIntoHtmlPages example finished." << std::endl << std::endl;
}