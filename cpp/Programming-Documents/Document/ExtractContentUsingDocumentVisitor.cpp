#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentVisitor.h>
#include <Model/Document/VisitorAction.h>
#include <Model/Fields/Nodes/FieldEnd.h>
#include <Model/Fields/Nodes/FieldSeparator.h>
#include <Model/Fields/Nodes/FieldStart.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/HeaderFooter.h>
#include <Model/Text/ControlChar.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/Run.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace
{
    class MyDocToTxtWriter : public DocumentVisitor
    {
        typedef MyDocToTxtWriter ThisType;
        typedef DocumentVisitor BaseType;

        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        MyDocToTxtWriter();
        System::String GetText();
        virtual VisitorAction VisitRun(System::SharedPtr<Run> run);
        virtual VisitorAction VisitFieldStart(System::SharedPtr<FieldStart> fieldStart);
        virtual VisitorAction VisitFieldSeparator(System::SharedPtr<FieldSeparator> fieldSeparator);
        virtual VisitorAction VisitFieldEnd(System::SharedPtr<FieldEnd> fieldEnd);
        virtual VisitorAction VisitParagraphEnd(System::SharedPtr<Paragraph> paragraph);
        virtual VisitorAction VisitBodyStart(System::SharedPtr<Body> body);
        virtual VisitorAction VisitBodyEnd(System::SharedPtr<Body> body);
        virtual VisitorAction VisitHeaderFooterStart(System::SharedPtr<HeaderFooter> headerFooter);

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        void AppendText(System::String text);

        System::SharedPtr<System::Text::StringBuilder> mBuilder;
        bool mIsSkipText;
    };

    RTTI_INFO_IMPL_HASH(425291245u, MyDocToTxtWriter, ThisTypeBaseTypesInfo);

    MyDocToTxtWriter::MyDocToTxtWriter() : mIsSkipText(false)
    {
        mIsSkipText = false;
        mBuilder = System::MakeObject<System::Text::StringBuilder>();
    }

    System::String MyDocToTxtWriter::GetText()
    {
        return mBuilder->ToString();
    }

    VisitorAction MyDocToTxtWriter::VisitRun(System::SharedPtr<Run> run)
    {
        AppendText(run->get_Text());
        // Let the visitor continue visiting other nodes.
        return VisitorAction::Continue;
    }

    VisitorAction MyDocToTxtWriter::VisitFieldStart(System::SharedPtr<FieldStart> fieldStart)
    {
        // In Microsoft Word, a field code (such as "MERGEFIELD FieldName") follows
        // After a field start character. We want to skip field codes and output field 
        // Result only, therefore we use a flag to suspend the output while inside a field code.
        // Note this is a very simplistic implementation and will not work very well
        // If you have nested fields in a document. 
        mIsSkipText = true;
        return VisitorAction::Continue;
    }

    VisitorAction MyDocToTxtWriter::VisitFieldSeparator(System::SharedPtr<FieldSeparator> fieldSeparator)
    {
        // Once reached a field separator node, we enable the output because we are
        // Now entering the field result nodes.
        mIsSkipText = false;
        return VisitorAction::Continue;
    }

    VisitorAction MyDocToTxtWriter::VisitFieldEnd(System::SharedPtr<FieldEnd> fieldEnd)
    {
        // Make sure we enable the output when reached a field end because some fields
        // Do not have field separator and do not have field result.
        mIsSkipText = false;
        return VisitorAction::Continue;
    }

    VisitorAction MyDocToTxtWriter::VisitParagraphEnd(System::SharedPtr<Paragraph> paragraph)
    {
        // When outputting to plain text we output Cr+Lf characters.
        AppendText(ControlChar::CrLf());
        return VisitorAction::Continue;
    }

    VisitorAction MyDocToTxtWriter::VisitBodyStart(System::SharedPtr<Body> body)
    {
        // We can detect beginning and end of all composite nodes such as Section, Body, 
        // Table, Paragraph etc and provide custom handling for them.
        mBuilder->Append(u"*** Body Started ***\r\n");
        return VisitorAction::Continue;
    }

    VisitorAction MyDocToTxtWriter::VisitBodyEnd(System::SharedPtr<Body> body)
    {
        mBuilder->Append(u"*** Body Ended ***\r\n");
        return VisitorAction::Continue;
    }

    VisitorAction MyDocToTxtWriter::VisitHeaderFooterStart(System::SharedPtr<HeaderFooter> headerFooter)
    {
        // Returning this value from a visitor method causes visiting of this
        // Node to stop and move on to visiting the next sibling node.
        // The net effect in this example is that the text of headers and footers
        // Is not included in the resulting output.
        return VisitorAction::SkipThisNode;
    }

    void MyDocToTxtWriter::AppendText(System::String text)
    {
        if (!mIsSkipText)
        {
            mBuilder->Append(text);
        }
    }

    System::Object::shared_members_type MyDocToTxtWriter::GetSharedMembers()
    {
        auto result = DocumentVisitor::GetSharedMembers();
        result.Add("MyDocToTxtWriter::mBuilder", this->mBuilder);
        return result;
    }
}

void ExtractContentUsingDocumentVisitor()
{
    std::cout << "ExtractContentUsingDocumentVisitor example started." << std::endl;
    // ExStart:ExtractContentUsingDocumentVisitor
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();

    // Open the document we want to convert.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Visitor.ToText.doc");

    // Create an object that inherits from the DocumentVisitor class.
    System::SharedPtr<MyDocToTxtWriter> myConverter = System::MakeObject<MyDocToTxtWriter>();

    // This is the well known Visitor pattern. Get the model to accept a visitor.
    // The model will iterate through itself by calling the corresponding methods
    // On the visitor object (this is called visiting).
    // Note that every node in the object model has the Accept method so the visiting
    // Can be executed not only for the whole document, but for any node in the document.
    doc->Accept(myConverter);

    // Once the visiting is complete, we can retrieve the result of the operation,
    // That in this example, has accumulated in the visitor.
    std::cout << myConverter->GetText().ToUtf8String() << std::endl;
    // ExEnd:ExtractContentUsingDocumentVisitor
    std::cout << "ExtractContentUsingDocumentVisitor example finished." << std::endl << std::endl;
}