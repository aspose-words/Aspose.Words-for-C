#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <system/date_time.h>
#include <system/collections/list.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class Comment; } }
namespace Aspose { namespace Words { class Document; } }
namespace Aspose { namespace Words { enum class VisitorAction; } }
namespace Aspose { namespace Words { class Run; } }
namespace Aspose { namespace Words { class CommentRangeStart; } }
namespace Aspose { namespace Words { class CommentRangeEnd; } }

namespace ApiExamples {

class ExComment : public ApiExampleBase
{
public:

    /// <summary>
    /// This Visitor implementation prints information about and contents of comments and comment ranges encountered in the document.
    /// </summary>
    class CommentInfoPrinter : public Aspose::Words::DocumentVisitor
    {
        typedef CommentInfoPrinter ThisType;
        typedef Aspose::Words::DocumentVisitor BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        CommentInfoPrinter();
        
        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        System::String GetText();
        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitRun(System::SharedPtr<Aspose::Words::Run> run) override;
        /// <summary>
        /// Called when a CommentRangeStart node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCommentRangeStart(System::SharedPtr<Aspose::Words::CommentRangeStart> commentRangeStart) override;
        /// <summary>
        /// Called when a CommentRangeEnd node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCommentRangeEnd(System::SharedPtr<Aspose::Words::CommentRangeEnd> commentRangeEnd) override;
        /// <summary>
        /// Called when a Comment node is encountered in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCommentStart(System::SharedPtr<Aspose::Words::Comment> comment) override;
        /// <summary>
        /// Called when the visiting of a Comment node is ended in the document.
        /// </summary>
        Aspose::Words::VisitorAction VisitCommentEnd(System::SharedPtr<Aspose::Words::Comment> comment) override;
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        bool mVisitorIsInsideComment;
        int32_t mDocTraversalDepth;
        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        
        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        void IndentAndAppendLine(System::String text);
        
    };
    
    
public:

    void AddCommentWithReply();
    void GetAllCommentsAndReplies();
    void RemoveCommentReplies();
    void RemoveCommentReply();
    void MarkCommentRepliesDone();
    void CreateCommentsAndPrintAllInfo();
    /// <summary>
    /// Create a new comment
    /// </summary>
    static System::SharedPtr<Aspose::Words::Comment> CreateComment(System::SharedPtr<Aspose::Words::Document> doc, System::String author, System::String initials, System::DateTime dateTime, System::String text);
    /// <summary>
    /// Extract comments from the document without replies.
    /// </summary>
    static System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Comment>>> ExtractComments(System::SharedPtr<Aspose::Words::Document> doc);
    
protected:

    /// <summary>
    /// Use an iterator and a visitor to print info of every comment from within a document.
    /// </summary>
    static void PrintAllCommentInfo(System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Comment>>> comments);
    
};

} // namespace ApiExamples


