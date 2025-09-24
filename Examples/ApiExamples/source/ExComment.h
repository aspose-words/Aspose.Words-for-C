#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeStart.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentVisitor.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExComment : public ApiExampleBase
{
    typedef ExComment ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Prints information and contents of all comments and comment ranges encountered in the document.
    /// </summary>
    class CommentInfoPrinter : public DocumentVisitor
    {
        typedef CommentInfoPrinter ThisType;
        typedef DocumentVisitor BaseType;
        
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
    void PrintAllComments();
    void RemoveCommentReplies();
    void Done();
    void CreateCommentsAndPrintAllInfo();
    void UtcDateTime();
    
protected:

    /// <summary>
    /// Iterates over every top-level comment and prints its comment range, contents, and replies.
    /// </summary>
    static void PrintAllCommentInfo(System::SharedPtr<Aspose::Words::NodeCollection> comments);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


