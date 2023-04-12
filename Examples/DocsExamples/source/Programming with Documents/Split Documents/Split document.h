#pragma once

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/Saving/DocumentSplitCriteria.h>
#include <Aspose.Words.Cpp/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <system/array.h>
#include <system/date_time.h>
#include <system/enum_helpers.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/io/directory.h>
#include <system/io/directory_info.h>
#include <system/io/file_system_info.h>
#include <system/io/search_option.h>
#include <system/linq/enumerable.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace DocsExamples { namespace Programming_with_Documents { namespace Split_Documents {

class SplitDocument : public DocsExamplesBase
{
public:
    void ByHeadingsHtml()
    {
        //ExStart:SplitDocumentByHeadingsHtml
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<HtmlSaveOptions>();
        options->set_DocumentSplitCriteria(DocumentSplitCriteria::HeadingParagraph);

        doc->Save(ArtifactsDir + u"SplitDocument.ByHeadingsHtml.html", options);
        //ExEnd:SplitDocumentByHeadingsHtml
    }

    void BySectionsHtml()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        //ExStart:SplitDocumentBySectionsHtml
        auto options = MakeObject<HtmlSaveOptions>();
        options->set_DocumentSplitCriteria(DocumentSplitCriteria::SectionBreak);
        //ExEnd:SplitDocumentBySectionsHtml

        doc->Save(ArtifactsDir + u"SplitDocument.BySectionsHtml.html", options);
    }

    void BySections()
    {
        //ExStart:SplitDocumentBySections
        auto doc = MakeObject<Document>(MyDir + u"Big document.docx");

        for (int i = 0; i < doc->get_Sections()->get_Count(); i++)
        {
            // Split a document into smaller parts, in this instance, split by section.
            SharedPtr<Section> section = doc->get_Sections()->idx_get(i)->Clone();

            auto newDoc = MakeObject<Document>();
            newDoc->get_Sections()->Clear();

            auto newSection = System::ExplicitCast<Section>(newDoc->ImportNode(section, true));
            newDoc->get_Sections()->Add(newSection);

            // Save each section as a separate document.
            newDoc->Save(ArtifactsDir + String::Format(u"SplitDocument.BySections_{0}.docx", i));
        }
        //ExEnd:SplitDocumentBySections
    }

    void PageByPage()
    {
        //ExStart:SplitDocumentPageByPage
        auto doc = MakeObject<Document>(MyDir + u"Big document.docx");

        int pageCount = doc->get_PageCount();

        for (int page = 0; page < pageCount; page++)
        {
            // Save each page as a separate document.
            SharedPtr<Document> extractedPage = doc->ExtractPages(page, 1);
            extractedPage->Save(ArtifactsDir + String::Format(u"SplitDocument.PageByPage_{0}.docx", page + 1));
        }
        //ExEnd:SplitDocumentPageByPage

        MergeDocuments();
    }

    //ExStart:MergeSplitDocuments
    void MergeDocuments()
    {
        // Find documents using for merge.
        ArrayPtr<SharedPtr<System::IO::FileSystemInfo>> documentPaths =
            MakeObject<System::IO::DirectoryInfo>(ArtifactsDir)
                ->GetFileSystemInfos(u"SplitDocument.PageByPage_*.docx")
                ->LINQ_OrderBy(System::Func<SharedPtr<System::IO::FileSystemInfo>, System::DateTime>([](SharedPtr<System::IO::FileSystemInfo> f)
                                                                                                     { return f->get_CreationTime(); }))
                ->LINQ_ToArray();
        String sourceDocumentPath =
            System::IO::Directory::GetFiles(ArtifactsDir, u"SplitDocument.PageByPage_1.docx", System::IO::SearchOption::TopDirectoryOnly)->idx_get(0);

        // Open the first part of the resulting document.
        auto sourceDoc = MakeObject<Document>(sourceDocumentPath);

        // Create a new resulting document.
        auto mergedDoc = MakeObject<Document>();
        auto mergedDocBuilder = MakeObject<DocumentBuilder>(mergedDoc);

        // Merge document parts one by one.
        for (SharedPtr<System::IO::FileSystemInfo> documentPath : documentPaths)
        {
            if (documentPath->get_FullName() == sourceDocumentPath)
            {
                continue;
            }

            mergedDocBuilder->MoveToDocumentEnd();
            mergedDocBuilder->InsertDocument(sourceDoc, ImportFormatMode::KeepSourceFormatting);
            sourceDoc = MakeObject<Document>(documentPath->get_FullName());
        }

        mergedDoc->Save(ArtifactsDir + u"SplitDocument.MergeDocuments.docx");
    }
    //ExEnd:MergeSplitDocuments

    void ByPageRange()
    {
        //ExStart:SplitDocumentByPageRange
        auto doc = MakeObject<Document>(MyDir + u"Big document.docx");

        // Get part of the document.
        SharedPtr<Document> extractedPages = doc->ExtractPages(3, 6);
        extractedPages->Save(ArtifactsDir + u"SplitDocument.ByPageRange.docx");
        //ExEnd:SplitDocumentByPageRange
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Split_Documents
