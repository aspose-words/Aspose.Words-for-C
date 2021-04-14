#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentSplitCriteria.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <system/date_time.h>
#include <system/enum_helpers.h>
#include <system/exceptions.h>
#include <system/io/directory_info.h>
#include <system/io/file_system_info.h>

#include "DocsExamplesBase.h"
#include "Programming with Documents/Split Documents/Page splitter.h"

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

            auto newSection = System::DynamicCast<Section>(newDoc->ImportNode(section, true));
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

        // Split nodes in the document into separate pages.
        auto splitter = MakeObject<DocumentPageSplitter>(doc);

        // Save each page as a separate document.
        for (int page = 1; page <= doc->get_PageCount(); page++)
        {
            SharedPtr<Document> pageDoc = splitter->GetDocumentOfPage(page);
            pageDoc->Save(ArtifactsDir + String::Format(u"SplitDocument.PageByPage_{0}.docx", page));
        }
        //ExEnd:SplitDocumentPageByPage

        MergeDocuments();
    }

    void ByPageRange()
    {
        //ExStart:SplitDocumentByPageRange
        auto doc = MakeObject<Document>(MyDir + u"Big document.docx");

        // Split nodes in the document into separate pages.
        auto splitter = MakeObject<DocumentPageSplitter>(doc);

        // Get part of the document.
        SharedPtr<Document> pageDoc = splitter->GetDocumentOfPageRange(3, 6);
        pageDoc->Save(ArtifactsDir + u"SplitDocument.ByPageRange.docx");
        //ExEnd:SplitDocumentByPageRange
    }

protected:
    void MergeDocuments()
    {
        // Find documents using for merge.
        ArrayPtr<SharedPtr<System::IO::FileSystemInfo>> documentPaths =
            MakeObject<System::IO::DirectoryInfo>(ArtifactsDir)
                ->GetFileSystemInfos(u"SplitDocument.PageByPage_*.docx")
                ->LINQ_OrderBy(static_cast<System::Func<SharedPtr<System::IO::FileSystemInfo>, System::DateTime>>(
                    static_cast<std::function<System::DateTime(SharedPtr<System::IO::FileSystemInfo> f)>>(
                        [](SharedPtr<System::IO::FileSystemInfo> f) -> System::DateTime { return f->get_CreationTime(); })))
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
};

}}} // namespace DocsExamples::Programming_with_Documents::Split_Documents
