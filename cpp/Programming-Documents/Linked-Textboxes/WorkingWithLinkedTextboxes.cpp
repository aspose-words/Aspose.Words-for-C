#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Drawing/Shape.h>
#include <Model/Drawing/ShapeType.h>
#include <Model/Drawing/TextBox.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace
{
    void CreateALink()
    {
        // ExStart:CreateALink
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<Shape> shape1 = System::MakeObject<Shape>(doc, ShapeType::TextBox);
        System::SharedPtr<Shape> shape2 = System::MakeObject<Shape>(doc, ShapeType::TextBox);

        System::SharedPtr<TextBox> textBox1 = shape1->get_TextBox();
        System::SharedPtr<TextBox> textBox2 = shape2->get_TextBox();

        if (textBox1->IsValidLinkTarget(textBox2))
        {
            textBox1->set_Next(textBox2);
        }
        // ExEnd:CreateALink
    }

    void CheckSequence()
    {
        // ExStart:CheckSequence
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<Shape> shape = System::MakeObject<Shape>(doc, ShapeType::TextBox);
        System::SharedPtr<TextBox> textBox = shape->get_TextBox();

        if ((textBox->get_Next() != nullptr) && (textBox->get_Previous() == nullptr))
        {
            std::cout << "The head of the sequence" << std::endl;
        }

        if ((textBox->get_Next() != nullptr) && (textBox->get_Previous() != nullptr))
        {
            std::cout << "The Middle of the sequence." << std::endl;
        }

        if ((textBox->get_Next() == nullptr) && (textBox->get_Previous() != nullptr))
        {
            std::cout << "The Tail of the sequence." << std::endl;
        }
        // ExEnd:CheckSequence
    }

    void BreakALink()
    {
        // ExStart:BreakALink
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<Shape> shape = System::MakeObject<Shape>(doc, ShapeType::TextBox);
        System::SharedPtr<TextBox> textBox = shape->get_TextBox();

        // Break a forward link
        textBox->BreakForwardLink();

        // Break a forward link by setting a null
        textBox->set_Next(nullptr);

        // Break a link, which leads to this textbox
        if (textBox->get_Previous() != nullptr)
        {
            textBox->get_Previous()->BreakForwardLink();
        }
        // ExEnd:BreakALink
    }
}

void WorkingWithLinkedTextboxes()
{
    std::cout << "WorkingWithLinkedTextboxes example started." << std::endl;
    CreateALink();
    CheckSequence();
    BreakALink();
    std::cout << "WorkingWithLinkedTextboxes example finished." << std::endl << std::endl;
}