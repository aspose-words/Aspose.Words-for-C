// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExVariableCollection.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/func.h>
#include <system/details/dispose_guard.h>
#include <system/collections/keyvalue_pair.h>
#include <system/collections/ienumerator.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldDocVariable.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/VariableCollection.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Fields;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2327351562u, ::Aspose::Words::ApiExamples::ExVariableCollection, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExVariableCollection : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExVariableCollection> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExVariableCollection>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExVariableCollection> ExVariableCollection::s_instance;

} // namespace gtest_test

void ExVariableCollection::Primer()
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
    //ExFor:VariableCollection.Item(Int32)
    //ExFor:VariableCollection.Item(String)
    //ExSummary:Shows how to work with a document's variable collection.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::VariableCollection> variables = doc->get_Variables();
    
    // Every document has a collection of key/value pair variables, which we can add items to.
    variables->Add(u"Home address", u"123 Main St.");
    variables->Add(u"City", u"London");
    variables->Add(u"Bedrooms", u"3");
    
    ASSERT_EQ(3, variables->get_Count());
    
    // We can display the values of variables in the document body using DOCVARIABLE fields.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldDocVariable>(builder->InsertField(Aspose::Words::Fields::FieldType::FieldDocVariable, true));
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
    ASSERT_TRUE(variables->LINQ_Any(static_cast<System::Func<System::Collections::Generic::KeyValuePair<System::String, System::String>, bool>>(static_cast<std::function<bool(System::Collections::Generic::KeyValuePair<System::String, System::String> v)>>([](System::Collections::Generic::KeyValuePair<System::String, System::String> v) -> bool
    {
        return v.get_Value() == u"London";
    }))));
    
    // The collection of variables automatically sorts variables alphabetically by name.
    ASSERT_EQ(0, variables->IndexOfKey(u"Bedrooms"));
    ASSERT_EQ(1, variables->IndexOfKey(u"City"));
    ASSERT_EQ(2, variables->IndexOfKey(u"Home address"));
    
    ASSERT_EQ(u"3", variables->idx_get(0));
    ASSERT_EQ(u"London", variables->idx_get(u"City"));
    
    // Enumerate over the collection of variables.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::Collections::Generic::KeyValuePair<System::String, System::String>>> enumerator = doc->get_Variables()->GetEnumerator();
        while (enumerator->MoveNext())
        {
            std::cout << System::String::Format(u"Name: {0}, Value: {1}", enumerator->get_Current().get_Key(), enumerator->get_Current().get_Value()) << std::endl;
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

namespace gtest_test
{

TEST_F(ExVariableCollection, Primer)
{
    s_instance->Primer();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
