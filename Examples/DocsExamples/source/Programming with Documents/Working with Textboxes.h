#pragma once

#include <iostream>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Drawing/TextBox.h>
#include <system/object.h>
#include <gtest/gtest.h>

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithTextboxes : public System::Object
{
public:
    void CreateALink()
    {
        //ExStart:CreateLink
        //GistId:e601abed5d181508ec5fed59ac29a14d
        auto doc = MakeObject<Document>();

        auto shape1 = MakeObject<Shape>(doc, ShapeType::TextBox);
        auto shape2 = MakeObject<Shape>(doc, ShapeType::TextBox);

        SharedPtr<TextBox> textBox1 = shape1->get_TextBox();
        SharedPtr<TextBox> textBox2 = shape2->get_TextBox();

        if (textBox1->IsValidLinkTarget(textBox2))
        {
            textBox1->set_Next(textBox2);
        }
        //ExEnd:CreateLink
    }

    void CheckSequence()
    {
        //ExStart:CheckSequence
        //GistId:e601abed5d181508ec5fed59ac29a14d
        auto doc = MakeObject<Document>();

        auto shape = MakeObject<Shape>(doc, ShapeType::TextBox);
        SharedPtr<TextBox> textBox = shape->get_TextBox();

        if (textBox->get_Next() != nullptr && textBox->get_Previous() == nullptr)
        {
            std::cout << "The head of the sequence" << std::endl;
        }

        if (textBox->get_Next() != nullptr && textBox->get_Previous() != nullptr)
        {
            std::cout << "The Middle of the sequence." << std::endl;
        }

        if (textBox->get_Next() == nullptr && textBox->get_Previous() != nullptr)
        {
            std::cout << "The Tail of the sequence." << std::endl;
        }
        //ExEnd:CheckSequence
    }

    void BreakALink()
    {
        //ExStart:BreakLink
        //GistId:e601abed5d181508ec5fed59ac29a14d
        auto doc = MakeObject<Document>();

        auto shape = MakeObject<Shape>(doc, ShapeType::TextBox);
        SharedPtr<TextBox> textBox = shape->get_TextBox();

        // Break a forward link.
        textBox->BreakForwardLink();

        // Break a forward link by setting a null.
        textBox->set_Next(nullptr);

        // Break a link, which leads to this textbox.
        if (textBox->get_Previous() != nullptr)
        {
            textBox->get_Previous()->BreakForwardLink();
        }
        //ExEnd:BreakLink
    }
};

}} // namespace DocsExamples::Programming_with_Documents
