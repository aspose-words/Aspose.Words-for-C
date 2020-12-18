#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/VariableCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldDocVariable.h>
#include <system/collections/ienumerator.h>
#include <system/collections/keyvalue_pair.h>
#include <system/details/dispose_guard.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace ApiExamples {

class ExVariableCollection : public ApiExampleBase
{
public:
    void Primer()
    {
        //ExStart
        //ExFor:Document.Variables
        //ExFor:VariableCollection
        //ExFor:VariableCollection.Add
        //ExFor:VariableCollection.Clear
        //ExFor:VariableCollection.Contains
        //ExFor:VariableCollection.Count
        //ExFor:VariableCollection.GetEnumerator
        //ExFor:VariableCollection.IndexOfKey
        //ExFor:VariableCollection.Remove
        //ExFor:VariableCollection.RemoveAt
        //ExSummary:Shows how to work with a document's variable collection.
        auto doc = MakeObject<Document>();
        SharedPtr<VariableCollection> variables = doc->get_Variables();

        // Documents have a variable collection to which name/value pairs can be added
        variables->Add(u"Home address", u"123 Main St.");
        variables->Add(u"City", u"London");
        variables->Add(u"Bedrooms", u"3");

        ASSERT_EQ(3, variables->get_Count());

        // Variables can be referenced and have their values presented in the document by DOCVARIABLE fields
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto field = System::DynamicCast<FieldDocVariable>(builder->InsertField(FieldType::FieldDocVariable, true));
        field->set_VariableName(u"Home address");
        field->Update();

        ASSERT_EQ(u"123 Main St.", field->get_Result());

        // Assigning values to existing keys will update them
        variables->Add(u"Home address", u"456 Queen St.");

        // DOCVARIABLE fields also need to be updated in order to show an accurate up to date value
        field->Update();

        ASSERT_EQ(u"456 Queen St.", field->get_Result());

        // The existence of variables can be looked up either by name or value like this
        ASSERT_TRUE(variables->Contains(u"City"));

        ASSERT_TRUE(variables->LINQ_Any([](auto v){ return v.get_Value() == u"London"; }));

        // Variables are automatically sorted in alphabetical order
        ASSERT_EQ(0, variables->IndexOfKey(u"Bedrooms"));
        ASSERT_EQ(1, variables->IndexOfKey(u"City"));
        ASSERT_EQ(2, variables->IndexOfKey(u"Home address"));

        // Enumerate over the collection of variables
        {
            SharedPtr<System::Collections::Generic::IEnumerator<System::Collections::Generic::KeyValuePair<String, String>>>
                enumerator = doc->get_Variables()->GetEnumerator();
            while (enumerator->MoveNext())
            {
                std::cout << "Name: " << enumerator->get_Current().get_Key() << ", Value: " << enumerator->get_Current().get_Value() << std::endl;
            }
        }

        // Variables can be removed either by name or index, or the entire collection can be cleared at once
        variables->Remove(u"City");

        ASSERT_FALSE(variables->Contains(u"City"));

        variables->RemoveAt(1);

        ASSERT_FALSE(variables->Contains(u"Home address"));

        variables->Clear();

        ASSERT_EQ(0, variables->get_Count());
        //ExEnd
    }
};

} // namespace ApiExamples
