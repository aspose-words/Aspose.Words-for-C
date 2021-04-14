#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IMailMergeDataSource.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <system/collections/list.h>
#include <system/io/directory.h>
#include <system/io/path.h>
#include <system/scope_guard.h>

#include "DocsExamplesBase.h"

namespace DocsExamples { namespace Programming_with_Documents { namespace Split_Documents {
class Topic;
}}} // namespace DocsExamples::Programming_with_Documents::Split_Documents

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace DocsExamples { namespace Programming_with_Documents { namespace Split_Documents {

class WordToHtmlConverter : public System::Object
{
private:
    class HandleTocMergeField : public IFieldMergingCallback
    {
    public:
        void FieldMerging(System::SharedPtr<FieldMergingArgs> e) override;
        void ImageFieldMerging(System::SharedPtr<ImageFieldMergingArgs> args) override;

    private:
        System::SharedPtr<DocumentBuilder> mBuilder;
    };

public:
    /// <summary>
    /// Performs the Word to HTML conversion.
    /// </summary>
    /// <param name="srcFileName">The MS Word file to convert.</param>
    /// <param name="tocTemplate">An MS Word file that is used as a template to build a table of contents.
    /// This file needs to have a mail merge region called "TOC" defined and one mail merge field called "TocEntry".</param>
    /// <param name="dstDir">The output directory where to write HTML files.</param>
    void Execute(System::String srcFileName, System::String tocTemplate, System::String dstDir);

private:
    System::SharedPtr<Document> mDoc;
    System::String mTocTemplate;
    System::String mDstDir;

    /// <summary>
    /// Selects heading paragraphs that must become topic starts.
    /// We can't modify them in this loop, so we need to remember them in an array first.
    /// </summary>
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Paragraph>>> SelectTopicStarts();
    /// <summary>
    /// Insert section breaks before the specified paragraphs.
    /// </summary>
    void InsertSectionBreaks(System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Paragraph>>> topicStartParas);
    /// <summary>
    /// Splits the current document into one topic per section and saves each topic
    /// as an HTML file. Returns a collection of Topic objects.
    /// </summary>
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Topic>>> SaveHtmlTopics();
    /// <summary>
    /// Leaves alphanumeric characters, replaces white space with underscore
    /// And removes all other characters from a string.
    /// </summary>
    System::String MakeTopicFileName(System::String paraText);
    /// <summary>
    /// Removes the last character (which is a paragraph break character from the given string).
    /// </summary>
    System::String MakeTopicTitle(System::String paraText);
    /// <summary>
    /// Saves one section of a document as an HTML file.
    /// Any embedded images are saved as separate files in the same folder as the HTML file.
    /// </summary>
    void SaveHtmlTopic(System::SharedPtr<Section> section, System::SharedPtr<Topic> topic);
    /// <summary>
    /// Generates a table of contents for the topics and saves to contents .html.
    /// </summary>
    void SaveTableOfContents(System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Topic>>> topics);
};

class SplitIntoHtmlPages : public DocsExamplesBase
{
public:
    void HtmlPages()
    {
        System::String srcFileName = MyDir + u"Footnotes and endnotes.docx";
        System::String tocTemplate = MyDir + u"Table of content template.docx";

        System::String outDir = System::IO::Path::Combine(ArtifactsDir, u"HtmlPages");
        System::IO::Directory::CreateDirectory_(outDir);

        auto w = System::MakeObject<WordToHtmlConverter>();
        w->Execute(srcFileName, tocTemplate, outDir);
    }
};

class Topic : public System::Object
{
public:
    System::String get_Title()
    {
        return pr_Title;
    }

    System::String get_FileName()
    {
        return pr_FileName;
    }

    Topic(System::String title, System::String fileName)
    {
        // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        System::Details::ThisProtector __local_self_ref(this);
        //---------------------------------------------------------Self Reference

        pr_Title = title;
        pr_FileName = fileName;
    }

private:
    System::String pr_Title;
    System::String pr_FileName;
};

class TocMailMergeDataSource : public IMailMergeDataSource
{
public:
    System::String get_TableName() override
    {
        return u"TOC";
    }

    TocMailMergeDataSource(System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Topic>>> topics) : mIndex(0)
    {
        mTopics = topics;
        mIndex = -1;
    }

    bool MoveNext() override
    {
        if (mIndex < mTopics->get_Count() - 1)
        {
            mIndex++;
            return true;
        }

        return false;
    }

    bool GetValue(System::String fieldName, System::SharedPtr<System::Object>& fieldValue) override
    {
        if (fieldName == u"TocEntry")
        {
            // The template document is supposed to have only one field called "TocEntry".
            fieldValue = mTopics->idx_get(mIndex);
            return true;
        }

        fieldValue.reset();
        return false;
    }

    System::SharedPtr<IMailMergeDataSource> GetChildDataSource(System::String tableName) override
    {
        ASPOSE_UNUSED(tableName);
        return nullptr;
    }

private:
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Topic>>> mTopics;
    int mIndex;
};

}}} // namespace DocsExamples::Programming_with_Documents::Split_Documents
