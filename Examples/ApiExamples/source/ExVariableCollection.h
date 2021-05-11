#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldDocVariable.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/VariableCollection.h>
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

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

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

        // Every document has a collection of key/value pair variables, which we can add items to.
        variables->Add(u"Home address", u"123 Main St.");
        variables->Add(u"City", u"London");
        variables->Add(u"Bedrooms", u"3");

        ASSERT_EQ(3, variables->get_Count());

        // We can display the values of variables in the document body using DOCVARIABLE fields.
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto field = System::DynamicCast<FieldDocVariable>(builder->InsertField(FieldType::FieldDocVariable, true));
        field->set_VariableName(u"Home address");
        field->Update();

        ASSERT_EQ(u"123 Main St.", field->get_Result());

        // Assigning values to existing keys will update them.
        variables->Add(u"Home address", u"456 Queen St.");

        // We will then have to update DOCVARIABLE fields to ensure they display an up-to-date value.
        ASSERT_EQ(u"123 Main St.", field->get_Result());

        field->Update();

        ASSERT_EQ(u"456 Queen St.", field->get_Result());

        // Verify that the document variables with a certain name or value exist.
        ASSERT_TRUE(variables->Contains(u"City"));
        ASSERT_TRUE(variables->LINQ_Any([](auto v) { return v.get_Value() == u"London"; }));

        // The collection of variables automatically sorts variables alphabetically by name.
        ASSERT_EQ(0, variables->IndexOfKey(u"Bedrooms"));
        ASSERT_EQ(1, variables->IndexOfKey(u"City"));
        ASSERT_EQ(2, variables->IndexOfKey(u"Home address"));

        // Enumerate over the collection of variables.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<System::Collections::Generic::KeyValuePair<String, String>>> enumerator =
                doc->get_Variables()->GetEnumerator();
            while (enumerator->MoveNext())
            {
                std::cout << "Name: " << enumerator->get_Current().get_Key() << ", Value: " << enumerator->get_Current().get_Value() << std::endl;
            }
        }

        // Below are three ways of removing document variables from a collection.
        // 1 -  By name:
        variables->Remove(u"City");

        ASSERT_FALSE(variables->Contains(u"City"));

        // 2 -  By index:
        variables->RemoveAt(1);

        ASSERT_FALSE(variables->Contains(u"Home address"));

        // 3 -  Clear the whole collection at once:
        variables->Clear();

        ASSERT_EQ(0, variables->get_Count());
        //ExEnd
    }
};

} // namespace ApiExamples
